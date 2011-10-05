'IMAP mailbox wrapper'
import re
import os
import gzip
import email
import random
import chardet
import imaplib
import datetime
import mimetypes
import logging; log = logging.getLogger(__name__)
from calendar import timegm
from email.generator import Generator
from email.parser import HeaderParser
from email.utils import mktime_tz, parsedate_tz, formatdate, parseaddr, formataddr, getaddresses
from email.header import decode_header, HeaderParseError


__all__ = ['IMAP4', 'IMAP4_SSL', 'IMAPError', 'Email', 'clean_nickname', 'parse_tags', 'format_tags', 'build_message', 'extract_parts']


pattern_folder = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?:\{.*\})?(?P<name>.*)')
pattern_whitespace = re.compile(r'\s+')
pattern_domain = re.compile(r'@[^,]+|/[^,]+')


class _IMAPExtension(object):
    'Mixin class that extends the IMAP interface'

    def __str__(self):
        return '%s@%s:%s' % (self.user, self.host, self.port)

    def format_error(self, text, data):
        'Format an error that happened with a server'
        return '[%s] %s%s' % (self, text, '\n' + repr(data))

    @property
    def folders(self):
        'Parse folder names'
        self.cd()
        folders = []
        r, data = self.list()
        if r != 'OK':
            raise IMAPError(self.format_error('Could not fetch folders', data))
        for item in data:
            if not item:
                continue
            if isinstance(item, tuple):
                item = ' '.join(item)
            folders.append(pattern_folder.match(item).groups()[2].lstrip())
        return folders

    def cd(self, folder=None):
        'Select the specified folder and return message count'
        r, data = self.select() if folder is None else self.select(folder)
        if r != 'OK':
            log.warn(self.format_error('[%s] Could not select folder' % format_tags(parse_tags(folder)), data))
            return 0
        return int(data[0])

    def walk(self, includes=None, excludes=None, searchCriterion=u'ALL', sortCriterion=u''):
        """
        Generate messages from folders matching includes/excludes and 
        messages matching search criterion ordered by sort criterion.

        Without arguments, it will walk all folders.
        With includes, it will walk only folders with matching names or matching parent names.
        With excludes, it will skip folders with matching names or matching parent names.
        With sortCriterion, it will return messages as sorted within a folder.
        Without sortCriterion, it will return messages in random order.

        Please see IMAP specification for more details on search and sort criteria.
        """
        # Prepare
        if sortCriterion:
            if 'SORT' not in self.capabilities:
                raise IMAPError(self.format_error('SORT not supported by server', self.capabilities))
            sortCriterion = '(%s)' % sortCriterion.encode('utf-8')
        searchCriterion = '(%s)' % searchCriterion.encode('utf-8') if searchCriterion else '(ALL)'
        if includes and not hasattr(includes, '__iter__'):
            includes = [includes]
        if excludes and not hasattr(excludes, '__iter__'):
            excludes = [excludes]
        includes = set(clean_tag(x) for x in includes) if includes else set()
        excludes = set(clean_tag(x) for x in excludes) if excludes else set()
        # Walk folders in random order
        folders = self.folders
        random.shuffle(folders)
        for folder in folders:
            # Read tags
            tags = parse_tags(folder)
            if excludes.intersection(tags):
                continue
            if includes and not includes.intersection(tags):
                continue
            # Get messageUIDs
            self.cd(folder)
            try:
                if sortCriterion:
                    r, data = self.uid('sort', sortCriterion, 'utf-8', searchCriterion)
                else:
                    r, data = self.uid('search', 'charset', 'utf-8', searchCriterion)
                if r != 'OK':
                    raise self.error(data)
            except self.error, error:
                log.warn(self.format_error("[%s] Could not execute searchCriterion=%s, sortCriterion=%s" % (format_tags(tags), searchCriterion, sortCriterion), error))
                continue
            messageUIDs = map(int, data[0].split())
            if not sortCriterion:
                random.shuffle(messageUIDs)
            # Walk messages
            for messageUID in messageUIDs:
                # Load message header
                try:
                    r, data = self.uid('fetch', messageUID, '(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM TO CC BCC DATE)])')
                    if r != 'OK':
                        raise self.error(data)
                # If we could not fetch the header, log it and move on
                except self.error, error:
                    log.warn(self.format_error('[%s UID=%s] Could not peek at message header' % (format_tags(tags), messageUID), error))
                    continue
                # Yield
                yield Email(self, messageUID, folder, data[0][1])

    def revive(self, targetFolder, message):
        'Upload the message to the targetFolder of the mail server'
        # Find the folder on the mail server
        for folder in self.folders:
            # If the folder exists, exit loop
            if folder.lower() == targetFolder.lower():
                break
        # If the folder does not exist, create it
        else:
            self.create(targetFolder)
            folder = targetFolder
        # A message without a date will return None instead of raising KeyError
        messageDate = message['date']
        r, data = self.append(folder, '', mktime_tz(parsedate_tz(messageDate)) if messageDate else None, message.as_string(False))
        if r != 'OK':
            raise IMAPError(self.format_error('Could not revive message', data))
        return data[0]


class IMAP4(_IMAPExtension, imaplib.IMAP4):
    'Extended IMAP4 client class'

    def __init__(self, host='', port=imaplib.IMAP4_PORT):
        self.host = host
        self.port = port
        imaplib.IMAP4.__init__(self, host, port)
        _IMAPExtension.__init__(self)

    def login(self, user, password):
        self.user = user
        imaplib.IMAP4.login(self, user, password)

    @classmethod
    def connect(cls, host='', port=None, user='', password=''):
        'Connect, login, return class instance'
        try:
            server = cls(host, port or imaplib.IMAP4_PORT)
            server.login(user, password)
        except Exception, error:
            server = _IMAPExtension()
            server.host = host
            server.port = port
            server.user = user
            raise IMAPError(server.format_error('Could not connect to server', error))
        return server


class IMAP4_SSL(_IMAPExtension, imaplib.IMAP4_SSL):
    'Extended IMAP4 client class over SSL connection'

    def __init__(self, host='', port=imaplib.IMAP4_SSL_PORT, keyfile=None, certfile=None):
        self.host = host
        self.port = port
        imaplib.IMAP4_SSL.__init__(self, host, port, keyfile, certfile)
        _IMAPExtension.__init__(self)

    def login(self, user, password):
        self.user = user
        imaplib.IMAP4_SSL.login(self, user, password)

    @classmethod
    def connect(cls, host='', port=None, user='', password='', keyfile=None, certfile=None):
        'Connect, login, return class instance'
        try:
            server = cls(host, port or imaplib.IMAP4_SSL_PORT, keyfile, certfile)
            server.login(user, password)
        except Exception, error:
            server = _IMAPExtension()
            server.host = host
            server.port = port
            server.user = user
            raise IMAPError(server.format_error('Could not connect to server', error))
        return server


class IMAPError(Exception):
    'IMAP error'
    pass


class Email(object):
    'Convenience class representing an email from an IMAP mailbox'

    def __init__(self, server, uid, folder, header):
        self.server = server
        self.uid = uid
        self.folder = folder
        self.tags = parse_tags(folder)
        # Parse header
        valueByKey = HeaderParser().parsestr(header)
        def getWhom(field):
            return ', '.join(formataddr((self._decode(x), y)) for x, y in getaddresses(valueByKey.get_all(field, [])))
        # Extract fields
        self.date = valueByKey.get('date')
        timePack = parsedate_tz(self.date)
        self.whenUTC = datetime.datetime.utcfromtimestamp(timegm(timePack) if timePack[-1] is None else mktime_tz(timePack)) if timePack else None
        self.subject = self._decode(valueByKey.get('subject', ''))
        self.fromWhom = getWhom('from')
        self.toWhom = getWhom('to')
        self.ccWhom = getWhom('cc')
        self.bccWhom = getWhom('bcc')

    def format_error(self, text, data):
        'Format an error that happened with a message'
        return self.server.format_error('[%s UID=%s] %s' % (format_tags(self.tags), self.uid, text), data)

    def _decode(self, text):
        'Decode text into utf-8'
        try:
            packs = decode_header(text)
        except HeaderParseError:
            try:
                packs = decode_header(text.replace('?==?', '?= =?'))
            except HeaderParseError:
                log.warn(self.format_error('Could not decode header', text))
                packs = [(text, 'utf-8')]
        string = ''.join(part.decode(encoding or 'utf-8', 'ignore') for part, encoding in packs)
        return pattern_whitespace.sub(' ', string.strip())

    @property
    def flags(self):
        'Get flags'
        r, data = self.server.uid('fetch', self.uid, '(FLAGS)')
        if r != 'OK':
            raise IMAPError(self.format_error('Could not get flags', data))
        string = data[0]
        return imaplib.ParseFlags(string) if data else ()

    @flags.setter
    def flags(self, flags):
        'Set flags'
        if not hasattr(flags, '__iter__'):
            flags = [flags]
        else:
            flags = list(flags)
        try:
            # Remove flags that we cannot set
            flags.remove(r'\Recent')
        except ValueError:
            pass
        r, data = self.server.uid('store', self.uid, 'FLAGS', '(%s)' % ' '.join(flags))
        if r != 'OK':
            raise IMAPError(self.format_error('Could not set flags', data))

    def set_flag(self, flag, on=True):
        'Set flag on or off'
        operator = '+' if on else '-'
        r, data = self.server.uid('store', self.uid, operator + 'FLAGS', '(%s)' % flag)
        if r != 'OK':
            raise IMAPError(self.format_error('Could not flag email', data))
        return self

    @property
    def seen(self):
        'Return True if email is marked as seen'
        return True if r'\Seen' in self.flags else False

    @seen.setter
    def seen(self, on=True):
        'Flag the email as seen or not'
        return self.set_flag(r'\Seen', on)

    @property
    def deleted(self):
        'Return True if email is marked as deleted'
        return True if r'\Deleted' in self.flags else False

    @deleted.setter
    def deleted(self, on=True):
        'Flag the email as deleted or not'
        return self.set_flag(r'\Deleted', on)

    def as_string(self, unixfrom=False):
        'Fetch mime string from server'
        try:
            # Get
            flags = self.flags
            # Load message
            r, data = self.server.uid('fetch', self.uid, '(RFC822)')
            if r != 'OK':
                raise IMAPError(self.format_error('Could not fetch message body', data))
            # Restore
            self.flags = flags
        except imaplib.IMAP4.abort, error:
            raise IMAPError(self.format_error('Connection failed while fetching message body', error))
        return data[0][1]

    def as_message(self):
        'Fetch mime string from server and convert to email.message.Message'
        return email.message_from_string(self.as_string())

    def save(self, targetPath=None):
        """
        Save email to the hard drive and return a list of parts by index, type, name.
        Compress the file if the filename ends with .gz
        Return partPacks if targetPath=None.
        """
        message = self.as_message()
        # Save
        if targetPath:
            Generator((gzip.open if targetPath.endswith('.gz') else open)(targetPath, 'wb')).flatten(message)
        # Gather partPacks
        partPacks = []
        for partIndex, part in enumerate(message.walk()):
            if 'multipart' == part.get_content_maintype():
                continue
            partPacks.append((partIndex, part.get_filename(), part.get_content_type()))
        return partPacks

    def __getitem__(self, key):
        return getattr(self, key)

    def __setitem__(self, key, value):
        return setattr(self, key, value)


def clean_nickname(text):
    'Extract nickname from email address'
    nickname, address = parseaddr(text)
    if not nickname:
        nickname = address
    return pattern_whitespace.sub(' ', pattern_domain.sub('', nickname).replace('.', ' ').replace('_', ' ')).strip('" ').title()


def clean_tag(text):
    'Convert to lowercase unicode, strip quotation marks, compact whitespace'
    return pattern_whitespace.sub(' ', text.lower().strip('" ')).decode('utf-8')


def parse_tags(text):
    'Parse tags from folder name'
    return [clean_tag(x) for x in text.replace('&-', '&').split('\\')]


def format_tags(tags, separator=' '):
    'Format tags'
    return separator.join(x.encode('utf-8') for x in tags)


def build_message(whenUTC=None, subject='', fromWhom='', toWhom='', ccWhom='', bccWhom='', bodyText='', bodyHTML='', attachmentPaths=None):
    'Build MIME message'
    mimeText = email.MIMEText.MIMEText(bodyText.encode('utf-8'), _charset='utf-8')
    mimeHTML = email.MIMEText.MIMEText(bodyHTML.encode('utf-8'), 'html')
    if attachmentPaths:
        message = email.MIMEMultipart.MIMEMultipart()
        if bodyText and bodyHTML:
            messageAlternative = email.MIMEMultipart.MIMEMultipart('alternative')
            messageAlternative.attach(mimeText)
            messageAlternative.attach(mimeHTML)
            message.attach(messageAlternative)
        elif bodyText:
            message.attach(mimeText)
        elif bodyHTML:
            message.attach(mimeHTML)
        for attachmentPath in attachmentPaths:
            attachmentName = os.path.basename(attachmentPath)
            contentType, contentEncoding = mimetypes.guess_type(attachmentName)
            # If we could not guess the type or the file is compressed,
            if contentType is None or contentEncoding is not None:
                contentType = 'application/octet-stream'
            mainType, subType = contentType.split('/', 1)
            payload = open(attachmentPath, 'rb').read()
            if mainType == 'text':
                part = email.MIMEText.MIMEText(payload, _subtype=subType, _charset=chardet.detect(payload)['encoding'])
            elif mainType == 'image':
                part = email.MIMEImage.MIMEImage(payload, _subtype=subType)
            elif mainType == 'audio':
                part = email.MIMEAudio.MIMEAudio(payload, _subtype=subType)
            else:
                part = email.MIMEBase.MIMEBase(mainType, subType)
                part.set_payload(payload)
                email.Encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=attachmentName)
            message.attach(part)
    elif bodyText and bodyHTML:
        message = email.MIMEMultipart.MIMEMultipart('alternative')
        message.attach(mimeText)
        message.attach(mimeHTML)
    elif bodyHTML:
        message = mimeHTML
    else:
        message = mimeText
    message['date'] = formatdate(timegm((whenUTC or datetime.datetime.utcnow()).timetuple()))
    message['subject'] = subject.encode('utf-8')
    message['from'] = fromWhom.encode('utf-8')
    message['to'] = toWhom.encode('utf-8')
    message['cc'] = ccWhom.encode('utf-8')
    message['bcc'] = bccWhom.encode('utf-8')
    return message


def extract_parts(sourcePath, partIndices=None, peek=False, applyCharset=True):
    """
    Get attachment from email sourcePath.
    Set partIndices=None to see all parts.
    Set peek=True to omit the payload.
    Set applyCharset=True to decode the payload into unicode.
    """
    if not hasattr(partIndices, '__iter__'):
        partIndices = [partIndices] if partIndices else []
    message = email.message_from_file(gzip.open(sourcePath, 'rb') if sourcePath.endswith('.gz') else open(sourcePath, 'rb'))
    partPacks = []
    for partIndex, part in enumerate(message.walk()):
        if 'multipart' == part.get_content_maintype():
            continue
        if partIndices and partIndex not in partIndices:
            continue
        partPack = partIndex, part.get_filename() or '', part.get_content_type() or ''
        if not peek:
            payload = part.get_payload(decode=True) or ''
            if applyCharset:
                charset = part.get_content_charset() or part.get_charset() or chardet.detect(payload)['encoding']
                payload = payload.decode(charset, 'ignore')
            partPack += (payload,)
        partPacks.append(partPack)
    return partPacks


def connect(host='', port=None, user='', password='', keyfile=None, certfile=None):
    'Connect to an IMAP server over an SSL connection'
    return IMAP4_SSL.connect(host, port, user, password, keyfile, certfile)
