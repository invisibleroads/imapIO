'IMAP mailbox wrapper'
import re
import os
import time
import random
import imaplib
import datetime
import mimetypes
import logging; log = logging.getLogger(__name__)
from email import message_from_string
from email.parser import HeaderParser
from email.header import decode_header, HeaderParseError
from email.utils import parsedate_tz


pattern_folder = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?:\{.*\})?(?P<name>.*)')
pattern_uid = re.compile(r'UID (\d+)')
pattern_whitespace = re.compile(r'\s+')


class _IMAPExtension(object):
    'Mixin class that extends the IMAP interface'

    def __str__(self):
        return '%s@%s:%s' % (self.user, self.host, self.port)

    def format_error(self, text, data=None):
        'Format error string'
        return '[%s] %s%s' % (self, text, '\n' + repr(data) if data else '')

    @property
    def folders(self):
        'Parse folder names'
        folders = []
        r, data = self.list()
        if r != 'OK':
            raise IMAPError(self.format_error('Could not fetch folders'))
        for item in data:
            if not item:
                continue
            if isinstance(item, tuple):
                item = ' '.join(item)
            folders.append(pattern_folder.match(item).groups()[2].lstrip())
        return folders

    def cycle(self, includes=None, excludes=None):
        """
        Generate messages from matching folders.
        Without arguments, it will cycle all folders.
        With includes, it will cycle only folders with matching names or matching parent names.
        With excludes, it will skip folders with matching names or matching parent names.
        """
        includes = {clean_tag(x) for x in includes} if includes else set()
        excludes = {clean_tag(x) for x in excludes} if excludes else set()
        # Cycle folders in random order
        folders = self.folders
        random.shuffle(folders)
        for folder in folders:
            # Extract tags
            tags = [clean_tag(x) for x in folder.replace('&-', '&').split('\\')]
            # If there are tags to exclude, skip the folder
            if excludes.intersection(tags):
                continue
            # If includes are defined and there are no tags to include, skip the folder
            if includes and includes.intersection(tags):
                continue
            # Cycle messages in random order
            messageCount = int(self.select(folder)[1][0])
            messageIndices = range(1, messageCount + 1)
            random.shuffle(messageIndices)
            for messageIndex in messageIndices:
                # Prepare error template
                format_error = lambda text, data: self.format_error('[INDEX=%s %s] %s' % (messageIndex, format_tags(tags), text), data)
                # Load message header
                try:
                    r, data = self.fetch(messageIndex, '(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM TO CC BCC DATE)] UID)')
                # If the connection died, log it and quit
                except imaplib.IMAP4.abort, error:
                    raise IMAPError(format_error('Connection failed while fetching message header', error))
                # If we could not fetch the header, log it and move on
                if r != 'OK':
                    log.warn(format_error('Could not peek at message header', data))
                    continue
                # Extract uid
                try:
                    match = pattern_uid.search(data[1])
                    if not match:
                        match = pattern_uid.search(data[0][0])
                        if not match:
                            raise IndexError
                # If we could not extract the uid, log it and move on
                except IndexError:
                    log.warn(format_error('Could not extract uid', data))
                    continue
                # Yield
                yield Message(self, int(match.group(1)), tags, data[0][1])

    def revive(self, folder, when, subject, fromWhom, toWhom, ccWhom, bccWhom, bodyText, bodyHtml, attachmentPaths):
        !!!
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
        # Append
        if when is None:
            when = message.when or datetime.datetime.now()
        r, data = self.append(folder, '', imaplib.Time2Internaldate(when.timetuple()), str(message))
        if r != 'OK':
            pass
        return 


class IMAP4(_IMAPExtension, imaplib.IMAP4):
    'Extended IMAP4 client class'

    def __init__(self, host='', port=143):
        self.host = host
        self.port = port
        imaplib.IMAP4.__init__(host, port)

    def login(self, user, password):
        self.user = user
        imaplib.IMAP4.login(user, password)


class IMAP4_SSL(_IMAPExtension, imaplib.IMAP4_SSL):
    'Extended IMAP4 client class over SSL connection'

    def __init__(self, host='', port=993, keyfile=None, certfile=None):
        self.host = host
        self.port = port
        imaplib.IMAP4.__init__(host, port, keyfile, certfile)

    def login(self, user, password):
        self.user = user
        imaplib.IMAP4.login(user, password)


class Message(object):
    'Convenience class representing an email message from an IMAP mailbox'

    def __init__(self, server, uid, tags, header):
        self.server = server
        self.uid = uid
        self.tags = tags
        # Parse header
        valueByKey = HeaderParser().parsestr(header)
        def getX(x):
            'Get value'
            if not x in valueByKey:
                return u''
            value = valueByKey[x]
            try:
                packs = decode_header(value)
            except HeaderParseError:
                try:
                    packs = decode_header(value.replace('?==?', '?= =?'))
                except HeaderParseError:
                    log.warn(self.format_error('Could not decode header', value))
                    packs = [(value, 'ascii')]
            string = ''.join(part.decode(encoding) for part, encoding in packs)
            return pattern_whitespace.sub(' ', string.strip())
        # Extract fields
        self.when = datetime.datetime.fromtimestamp(time.mktime(parsedate_tz(message['Date'])[:9])) if 'Date' in message else None
        self.subject = getX('Subject')
        self.fromWhom = getX('From')
        self.toWhom = getX('To')
        self.ccWhom = getX('CC')
        self.bccWhom = getX('BCC')

    def format_error(self, text, data=None):
        'Format error string'
        return self.server.format_error('[UID=%s %s] %s' % (self.uid, format_tags(self.tags), text), data)

    @property
    def flags(self):
        'Get message flags'
        r, data = self.server.uid('fetch', self.uid, '(FLAGS)')
        if r != 'OK':
            raise IMAPError(self.format_error('Could not get flags', data))
        string = data[0]
        return imaplib.ParseFlags(string) if data else ()

    @flags.setter
    def flags(self, flags):
        r, data = self.server.uid('store', self.uid, 'FLAGS', '(%s)' % ' '.join(flags))
        if r != 'OK':
            raise IMAPError(self.format_error('Could not set flags', data))

    @property
    def is_unread(self):
        'Return True if message is marked as unread'
            return False
        return False if r'\Seen' in self.flags else True

    def mark_unread(self):
        'Mark message as unread'
        r, data = self.server.uid('store', self.uid, '-FLAGS', r'(\Seen)')
        if r != 'OK':
            raise IMAPError(self.format_error('Could not mark message as unread', data))
        return self

    def mark_deleted(self):
        'Mark message as deleted'
        r, data = self.server.uid('store', self.uid, '+FLAGS', r'(\Deleted)')
        if r != 'OK':
            raise IMAPError(self.format_error('Could not mark message as deleted', data))
        return self

    def download(self, targetFolderPath):
        'Save message and its attachments to the hard drive'
        try:
            # Save flags
            flags = self.flags
            # Load message
            r, data = self.server.uid('fetch', self.uid, '(RFC822)')
            if r != 'OK':
                raise IMAPError(self.format_error('Could not fetch message body', data))
            msg = message_from_string(data[0][1])
            # Restore flags
            self.flags = flags
        except imaplib.IMAP4.abort, error:
            raise IMAPError(self.format_error('Connection failed while fetching message body', error))
        # Save parts using a coroutine
        with open(os.path.join(targetFolderPath, 'parts.txt'), 'wt') as partsFile:
            partConsumer = _make_partConsumer(targetFolderPath, partsFile)
            partConsumer.next()
            for part in msg.walk():
                # If the content is multipart, then enter the container
                if 'multipart' == part.get_content_maintype():
                    continue
                # Consume part
                partConsumer.send(part)
        # Save tags and header
        open(os.path.join(targetFolderPath, 'tags.txt'), 'wt').write(format_tags(self.tags, '\n'))
        open(os.path.join(targetFolderPath, 'header.txt'), 'wt').write(format_header(self.when, self.subject, self.fromWhom, self.toWhom, self.ccWhom, self.bccWhom))

    def __getitem__(self, key):
        return getattr(self, key)

    def __setitem__(self, key, value):
        return setattr(self, key, value)


class IMAPError(Exception):
    pass


def _make_partConsumer(targetFolderPath, partsFile):
    'Make a coroutine that saves the parts of an email message to the hard drive'
    index = 1
    while True:
        part = (yield)
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        contentType = part.get_content_type().lower()
        fileName = part.get_filename()



        if contentType in ['text/plain', 'text/html']:
            payload = payload.decode(part.get_content_charset() or charset.detect(payload)['encoding']).encode('utf8')
            extension = mimetypes.guess_extension(contentType)
        else:
            extension = '.bin'



            open(os.path.join(targetFolderPath, partName + extension)).write(payload.encode('utf8'))
        else:
            open(os.path.join(targetFolderPath, partName



            
        targetPath = os.path.join(targetFolderPath, 'part%d%s' % (index, extension)

        # Encode decode?
        open(os.path.join(targetFolderPath, partName), 'wb').write(payload)
        



        # Increment
        count += 1


                else:
                    extension = mimetypes.guess_extension(contentType)
                    # If we could not guess an extension,
                    if not extension:
                        # Use generic extension
                        extension = '.bin'
                filename = 'part%03d%s' % (counter, extension)

            fp = open(os.path.join(targetFolderPath, filename), 'wb')
            fp.write(payload)
            fp.close()
        # Return
        return True


pattern_address = re.compile(r'<.*?>')
pattern_bracket = re.compile(r'<|>')
pattern_domain = re.compile(r'@[^,]+|/[^,]+')


def format_header(when, subject, fromWhom, toWhom, ccWhom, bccWhom):
    pass
def formatHeader(subject, when, fromWhom, toWhom, ccWhom, bccWhom):
    # Unicode everything
    subject = unicodeSafely(subject)
    fromWhom = unicodeSafely(fromWhom)
    toWhom = unicodeSafely(toWhom)
    ccWhom = unicodeSafely(ccWhom)
    bccWhom = unicodeSafely(bccWhom)
    # Build header
    header = 'From:       %(fromWhom)s\nDate:       %(date)s\nSubject:    %(subject)s' % {
        'fromWhom': fromWhom,
        'date': when.strftime('%A, %B %d, %Y    %I:%M:%S %p'),
        'subject': subject,
    }
    # Add optional features
    if toWhom:
        header += '\nTo:         %s' % toWhom
    if ccWhom:
        header += '\nCC:         %s' % ccWhom
    if bccWhom:
        header += '\nBCC:        %s' % bccWhom
    # Return
    return header


def format_to_cc_bcc:
    pass

def formatToWhomString(toWhom, ccWhom, bccWhom):
    return ', '.join(filter(lambda x: x, (y.strip() for y in (toWhom, ccWhom, bccWhom))))


def clean_nickname(text):
    pass
def formatNameString(nameString):
    # Split the string
    oldTerms = nameString.split(',')
    newTerms = []
    # For each term,
    for oldTerm in oldTerms:
        # Try removing the address
        newTerm = pattern_address.sub('', oldTerm).strip()
        # If the term is empty,
        if not newTerm:
            # Remove brackets
            newTerm = pattern_bracket.sub('', oldTerm).strip()
        # Append
        newTerms.append(newTerm)
    # Remove domain
    string = pattern_domain.sub('', ', '.join(newTerms))
    # Return
    return string.replace('"', '').replace('.', ' ')


def clean_filename(text):
    pass
def sanitizeFileName(fileName):
    alphabet = "-_.() %s%s" % (string.ascii_letters, string.digits)
    return ''.join(x if x in alphabet else '-' for x in fileName)


def clean_tag(text):
    'Convert to lowercase unicode, strip quotation marks, compact whitespace'
    return pattern_whitespace.sub(' ', text.lower().strip('" ')).decode('utf8')


def buildMessage(headerByValue, when, bodyText='', bodyHtml='', attachmentPaths=None):
# Build
message = email.MIMEMultipart.MIMEMultipart()
for key, value in headerByValue.iteritems():
message[key] = value
message['Date'] = email.utils.formatdate(time.mktime(when.timetuple()), localtime=True)

# Set body
mimeText = email.MIMEText.MIMEText(bodyText)
mimeBody = email.MIMEText.MIMEText(bodyHtml, 'html')
if bodyText and bodyHtml:
messageAlternative = email.MIMEMultipart.MIMEMultipart('alternative')
messageAlternative.attach(mimeText)
messageAlternative.attach(mimeBody)
message.attach(messageAlternative)
elif bodyText:
message.attach(mimeText)
elif bodyHtml:
message.attach(mimeBody)

# Set attachments
if attachmentPaths:
for attachmentPath in attachmentPaths:
attachmentName = os.path.basename(attachmentPath)
mimeType = mimetypes.guess_type(attachmentName)[0]
if not mimeType:
mimeType = 'application/octet-stream'
part = email.MIMEBase.MIMEBase(*mimeType.split('/'))
part.set_payload(open(attachmentPath, 'rb').read())
email.Encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="%s"' % attachmentName)
message.attach(part)
# Return
return message


def format_tags(tags, separator=' '):
    'Format tags'
    return separator.join(x.encode('utf8') for x in tags)
