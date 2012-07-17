'IMAP mailbox wrapper'
import chardet
import datetime
import email
import gzip
import imaplib
import logging; log = logging.getLogger(__name__)
import mimetypes
import os
import random
import re
from calendar import timegm
from email.generator import Generator
from email.header import decode_header, HeaderParseError
from email.parser import HeaderParser
from email.utils import mktime_tz, parsedate_tz, formatdate, parseaddr, formataddr, getaddresses

from imapIO import utf_7_imap4


__all__ = ['IMAP4', 'IMAP4_SSL', 'IMAPError', 'Email', 'build_message', 'normalize_nickname', 'connect', 'extract']


PATTERN_FOLDER = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?:\{.*\})?(?P<name>.*)')
PATTERN_WHITESPACE = re.compile(r'\s+')
PATTERN_DOMAIN = re.compile(r'@[^,]+|/[^,]+')


class _IMAPExtension(object):
    'Mixin class that extends the IMAP interface'

    host = ''

    def __init__(self):
        if 'imap.mail.yahoo.com' == self.host.lower():
            self.xatom('ID ("GUID" "1")')

    def __str__(self):
        return '%s:%s %s' % (self.host, self.port, self.user)

    def format_error(self, text, data):
        'Format an error that happened with a server'
        return '[%s]\n%s\n%s' % (self, text, str(data))

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
            if hasattr(item, '__iter__'):
                item = ' '.join(item)
            folders.append(PATTERN_FOLDER.match(item).groups()[2].lstrip())
        return folders

    def cd(self, folder=None):
        'Select the specified folder and return message count'
        r, data = self.select() if folder is None else self.select(folder)
        if r != 'OK':
            log.warn(self.format_error('[%s] Could not select folder' % folder, data))
            return 0
        return int(data[0])

    def walk(self, include=lambda folder: True, searchCriterion=u'ALL', sortCriterion=u'', shuffleMessages=True):
        """
        Yield matching messages from matching folders.
        Without arguments, it will yield messages in random order.
        Specify a folder, a list of folders or a function as the first argument.
        See IMAP specification for details on search and sort criteria.

        Yield messages from folders that start with the letter A.
            server.walk(lambda folder: folder.upper().startswith('A'))

        Yield messages from non-trash folders.
            server.walk(lambda folder: folder.lower() not in ['trash', 'spam'])
        """
        include = make_folderFilter(include)
        searchCriterion = '(%s)' % searchCriterion.encode('utf-8')
        if sortCriterion:
            if 'SORT' not in self.capabilities:
                raise IMAPError(self.format_error('SORT not supported by server', self.capabilities))
            sortCriterion = '(%s)' % sortCriterion.encode('utf-8')
        # Walk folders
        folders = self.folders
        random.shuffle(folders)
        for folder in folders:
            if not include(folder):
                continue
            self.cd(folder)
            try:
                if sortCriterion:
                    r, data = self.uid('sort', sortCriterion, 'utf-8', searchCriterion)
                else:
                    r, data = self.uid('search', 'charset', 'utf-8', searchCriterion)
                if r != 'OK':
                    raise self.error(data)
            except self.error, error:
                log.warn(self.format_error("[%s] Could not load messageUIDs" % folder, error))
                continue
            messageUIDs = [int(x) for x in data[0].split()]
            if shuffleMessages and not sortCriterion:
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
                    log.warn(self.format_error('[%s UID=%s] Could not peek at message header' % (folder, messageUID), error))
                    continue
                yield Email(self, messageUID, folder, data[0][1])

    def revive(self, targetFolder, message):
        'Upload the message to the targetFolder of the mail server'
        # Find the folder on the mail server
        for folder in self.folders:
            # If the folder exists, exit loop
            if normalize_folder(folder) == normalize_folder(targetFolder):
                break
        # If the folder does not exist, create it
        else:
            self.create(targetFolder)
            folder = targetFolder
        # A message with no date returns None instead of raising KeyError
        messageDate = message['date']
        r, data = self.append(folder, '', mktime_tz(parsedate_tz(messageDate)) if messageDate else None, message.as_string(False))
        if r != 'OK':
            raise IMAPError(self.format_error('Could not revive message', data))
        return data[0]


class IMAP4(_IMAPExtension, imaplib.IMAP4): # pragma: no cover
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


class IMAP4_SSL(_IMAPExtension, imaplib.IMAP4_SSL): # pragma: no cover
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
        self.header = header
        # Parse header
        valueByKey = HeaderParser().parsestr(header)
        def getWhom(field):
            return ', '.join(formataddr((self._decode(x), self._decode(y))) for x, y in getaddresses(valueByKey.get_all(field, [])))
        # Extract fields
        self.date = valueByKey.get('date')
        timePack = parsedate_tz(self.date)
        if not timePack:
            self.whenUTC = None
            self.whenLocal = None
        else:
            timeStamp = timegm(timePack) if timePack[-1] is None else mktime_tz(timePack)
            self.whenUTC = datetime.datetime.utcfromtimestamp(timeStamp)
            self.whenLocal = datetime.datetime.fromtimestamp(timeStamp)
        self.subject = self._decode(valueByKey.get('subject', ''))
        self.fromWhom = getWhom('from')
        self.toWhom = getWhom('to')
        self.ccWhom = getWhom('cc')
        self.bccWhom = getWhom('bcc')

    def __getitem__(self, key):
        keyLower = key.lower()
        if keyLower in ['from', 'to', 'cc', 'bcc']:
            return getattr(self, keyLower + 'Whom')
        return getattr(self, key)

    def __setitem__(self, key, value):
        keyLower = key.lower()
        if keyLower in ['from', 'to', 'cc', 'bcc']:
            return setattr(self, keyLower + 'Whom', value)
        return setattr(self, key, value)

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
        return PATTERN_WHITESPACE.sub(' ', string.strip())

    def format_error(self, text, data):
        'Format an error that happened with a message'
        return self.server.format_error('[%s UID=%s] %s' % (self.folder, self.uid, text), data)

    @property
    def flags(self):
        'Get flags'
        r, data = self.server.uid('fetch', self.uid, '(FLAGS)')
        if r != 'OK':
            raise IMAPError(self.format_error('Could not get flags', data))
        string = data[0]
        return imaplib.ParseFlags(string) if string else ()

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
        if not hasattr(self, '_string'):
            try:
                # Get
                flags = self.flags
                # Load message
                r, data = self.server.uid('fetch', self.uid, '(RFC822)')
                if r != 'OK':
                    raise IMAPError(self.format_error('Could not fetch body', data))
                # Restore
                self.flags = flags
            except imaplib.IMAP4.abort, error:
                message = 'Connection failed while fetching body'
                raise IMAPError(self.format_error(message, error))
            self._string = data[0][1]
        return self._string

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
            mainType = part.get_content_maintype()
            if 'multipart' == mainType:
                continue
            partName = part.get_filename() or ''
            partType = part.get_content_type() or ''
            partPack = partIndex, partName, partType
            partPacks.append(partPack)
        return partPacks

    def extract(self, include=lambda index, name, type: True, peek=False, applyCharset=True):
        return extract(self.as_message(), include, peek, applyCharset)


def build_message(whenUTC=None, subject='', fromWhom='', toWhom='', ccWhom='', bccWhom='', bodyText='', bodyHTML='', attachmentPaths=None):
    'Build MIME message'
    subject, bodyText, bodyHTML = map(strip_illegal_characters, [subject, bodyText, bodyHTML])
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


def connect(host='', port=None, user='', password='', keyfile=None, certfile=None):
    'Connect to an IMAP server over an SSL connection'
    return IMAP4_SSL.connect(host, port, user, password, keyfile, certfile)


def extract(source, include=lambda index, name, type: True, peek=False, applyCharset=True):
    """
    Get message parts, where source is either an instance of 
    MIMEMessage or a path to a file containing a MIMEMessage.

    Set include=lambda index, name, type: type.startswith('image') to extract images only.
    Set peek=True to omit the payload.
    Set applyCharset=True to decode the payload into unicode.
    """
    if hasattr(source, 'walk'):
        message = source
    else:
        sourceFile = gzip.open(source, 'rb') if source.endswith('.gz') else open(source, 'rb')
        message = email.message_from_file(sourceFile)
    partPacks = []
    for partIndex, part in enumerate(message.walk()):
        mainType = part.get_content_maintype()
        if 'multipart' == mainType:
            continue
        partName = part.get_filename() or ''
        partType = part.get_content_type() or ''
        partPack = partIndex, partName, partType
        if not include(*partPack):
            continue
        if not peek:
            payload = part.get_payload(decode=True) or ''
            if 'text' == mainType and applyCharset:
                charset = part.get_content_charset() or part.get_charset() or chardet.detect(payload)['encoding']
                payload = payload.decode(charset, 'ignore')
            partPack += (payload,)
        partPacks.append(partPack)
    return partPacks


def make_folderFilter(x):
    # If x is unicode or a string,
    if hasattr(x, 'lower'):
        x = normalize_folder(x)
        return lambda folder: normalize_folder(folder) == x
    # If x is a tuple or a list or a dictionary,
    if hasattr(x, '__iter__'):
        x = [normalize_folder(y) for y in x]
        return lambda folder: normalize_folder(folder) in x
    # If x is a function,
    if hasattr(x, 'func_code'):
        return lambda folder: x(normalize_folder(folder))


def normalize_folder(text):
    text = text.decode('utf-7-imap4')
    text = text.strip('" ')
    text = text.lower()
    return PATTERN_WHITESPACE.sub(' ', text)


def normalize_nickname(text):
    'Extract nickname from email address'
    nickname, address = parseaddr(text)
    if not nickname:
        nickname = address
    return PATTERN_WHITESPACE.sub(' ', PATTERN_DOMAIN.sub('', nickname).replace('.', ' ').replace('_', ' ')).strip('" ').title()


def strip_illegal_characters(x):
    return x.replace(chr(0), '')
