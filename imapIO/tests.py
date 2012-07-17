# -*- coding: utf-8 -*-
'Tests for imapIO'
import os
import random
import tempfile
import unittest
import datetime
import ConfigParser
import logging; logging.basicConfig()

import imapIO
from imapIO.utf_7_imap4 import CODEC_NAME


configuration = ConfigParser.ConfigParser()
if not configuration.read('.test.ini'):
    raise Exception('Please create a configuration file called .test.ini')
getX = lambda x: configuration.get('imap', x)
host = getX('host')
port = int(getX('port'))
user = getX('user')
password = getX('password')
ssl = getX('ssl').lower() == 'true'


class Base(object):

    def tearDown(self):
        if hasattr(self, 'temporaryPaths'):
            for x in self.temporaryPaths:
                if os.path.exists(x):
                    os.remove(x)

    def test_folders(self):
        self.server.folders

    def test_cd(self):
        self.server.cd(random.choice(self.server.folders))
        self.server.cd()

    def test_walk(self):
        # Test include
        folder = random.choice(self.server.folders)
        for include in folder, [folder], lambda x: x.lower() == folder.lower():
            try:
                email = self.server.walk(include).next()
            except StopIteration:
                pass
            else:
                self.assertEqual(True, imapIO.make_folderFilter(include)(email.folder))
        # Test searchCriterion
        try:
            self.server.walk(searchCriterion=u'SINCE 01-JAN-2006 BEFORE 01-JAN-2007').next()
        except StopIteration:
            pass
        # Test sortCriterion
        if 'SORT' in self.server.capabilities:
            try:
                self.server.walk(sortCriterion='ARRIVAL').next()
            except StopIteration:
                pass

    def test_revive(self):
        folder = 'inbox'
        self.server.cd(folder)
        baseCase = dict(
            whenUTC=datetime.datetime(2005, 1, 23, 1, 0),
            subject='Test',
            fromWhom='from@example.com',
            toWhom='to@example.com',
            ccWhom='cc@example.com',
            bccWhom='bcc@example.com',
            bodyText='Yes',
            bodyHTML='<html>No</html>',
            attachmentPaths=[
                'CHANGES.rst',
                'README.rst',
            ])
        # Clear previous cases
        for email in self.server.walk(folder):
            if [email.fromWhom, email.toWhom] == [baseCase['fromWhom'], baseCase['toWhom']]:
                email.deleted = True
        self.server.expunge()
        # Run cases
        cases = [
            baseCase,
            dict(baseCase, bodyHTML=''),
            dict(baseCase, bodyText=''),
            dict(baseCase, attachmentPaths=None),
            dict(baseCase, attachmentPaths=None, bodyHTML=''),
            dict(baseCase, attachmentPaths=None, bodyText=''),
        ]
        self.temporaryPaths = []
        for caseIndex, case in enumerate(cases):
            subject = case['subject'] + str(caseIndex)
            # Revive
            self.server.revive(folder, imapIO.build_message(**dict(case, subject=subject)))
            # Make sure the revived email exists
            for email in self.server.walk(folder):
                if [email.fromWhom, email.toWhom, email.subject] == [baseCase['fromWhom'], baseCase['toWhom'], subject]:
                    break
            else:
                raise AssertionError('Could not find revived message on server')
            self.assertEqual(email.seen, False)
            self.assertEqual(email.whenUTC, case['whenUTC'])
            email.flags = r'\Seen'
            self.assertEqual(
                set([r'\Seen']), 
                set(x for x in email.flags).difference([r'\Recent']))
            # Save
            targetPath = tempfile.mkstemp(suffix='.gz')[1]
            self.temporaryPaths.append(targetPath)
            email.extract(lambda index, name, type: 'text/html' == type, peek=True)
            email.save(targetPath)
            partPacks = imapIO.extract(targetPath)
            attachmentPathByName = dict((os.path.basename(x), x) for x in case['attachmentPaths'] or [])
            # Make sure the email contains all attachments
            self.assertEqual(set(attachmentPathByName) - set(x[1] for x in partPacks), set())
            # Make sure attachment contents match
            for partIndex, partName, contentType, payload in partPacks:
                if partName in attachmentPathByName:
                    attachmentData = open(attachmentPathByName[partName], 'rb').read()
                    if contentType.startswith('text'):
                        payload = payload.replace('\r\n', '\n')
                    self.assertEqual(payload, attachmentData)
                elif contentType == 'text/plain':
                    self.assertEqual(payload, case['bodyText'])
                elif contentType == 'text/html':
                    self.assertEqual(payload, case['bodyHTML'])
                else:
                    raise Exception('Unexpect part: %s' % (partIndex, partName, contentType))
        # Duplicate an email directly
        for email in self.server.walk(folder):
            if [email.fromWhom, email.toWhom] == [baseCase['fromWhom'], baseCase['toWhom']]:
                self.server.revive(folder, email)
                break
        # Clear cases
        self.server.format_error('xxx', '')
        for email in self.server.walk(folder):
            email.format_error('xxx', '')
            if [email.fromWhom, email.toWhom] == [baseCase['fromWhom'], baseCase['toWhom']]:
                email.deleted = True
        self.server.expunge()


@unittest.skipIf(ssl, 'not configured')
class TestIMAP4(unittest.TestCase, Base):

    def setUp(self):
        if not hasattr(self, 'server'):
            self.server = imapIO.IMAP4.connect(host, port, user, password)


@unittest.skipIf(not ssl, 'not configured')
class TestIMAP4_SSL(unittest.TestCase, Base):

    def setUp(self):
        if not hasattr(self, 'server'):
            self.server = imapIO.connect(host, port, user, password)


class TestExceptions_IMAPExtension(unittest.TestCase):

    def setUp(self):
        self.server = IMAP4Dummy()

    def test_folders(self):
        with self.assertRaises(imapIO.IMAPError):
            self.server.cd = lambda: None
            self.server.list = lambda: ('xxx', [])
            self.server.folders
        self.server.list = lambda: ('OK', [''])
        self.server.folders
        self.server.list = lambda: ('OK', [('()', '"/"', 'xxx')])
        self.server.folders

    def test_cd(self):
        self.server.select = lambda: ('xxx', [])
        self.server.cd()

    def test_walk(self):
        self.server.capabilities = []
        with self.assertRaises(imapIO.IMAPError):
            self.server.walk(sortCriterion='ARRIVAL').next()
        self.server.cd = lambda a='': None
        self.server.list = lambda: ('OK', ['() "/" aaa', '() "/" bbb'])
        self.server.uid = lambda a, b, c, d: ('xxx', [])
        with self.assertRaises(StopIteration):
            self.server.walk('bbb').next()
        self.server.uid = lambda a, b, c, d=None: ('OK' if a.startswith('s') else 'xxx', ['1'])
        with self.assertRaises(StopIteration):
            self.server.walk('bbb').next()

    def test_revive(self):
        self.server.cd = lambda a='': None
        self.server.list = lambda: ('OK', ['() "/" aaa'])
        self.server.create = lambda a: None
        self.server.append = lambda a, b, c, d: ('xxx', [])
        with self.assertRaises(imapIO.IMAPError):
            self.server.revive('bbb', imapIO.build_message())


class TestExceptions_Email(unittest.TestCase):

    def setUp(self):
        self.server = IMAP4Dummy()
        self.email = imapIO.Email(self.server, None, '', '')

    def test_decode(self):
        decode_header = imapIO.decode_header
        def raise_exception(a):
            raise imapIO.HeaderParseError
        imapIO.decode_header = raise_exception
        imapIO.Email(self.server, None, '', '')
        imapIO.decode_header = decode_header

    def test_flags(self):
        self.server.uid = lambda a, b, c, d=None: ('xxx', [])
        with self.assertRaises(imapIO.IMAPError):
            self.email.flags
        with self.assertRaises(imapIO.IMAPError):
            self.email.flags = ''
        with self.assertRaises(imapIO.IMAPError):
            self.email.seen = True
        self.server.uid = lambda a, b, c, d=None: ('OK', [''])
        self.email.deleted

    def test_as_string(self):
        self.server.uid = lambda a, b, c, d=None: ('OK' if c == '(FLAGS)' else 'xxx', ['1'])
        with self.assertRaises(imapIO.IMAPError):
            self.email.as_string()
        def raise_exception(a, b, c):
            raise imapIO.imaplib.IMAP4.abort
        self.server.uid = raise_exception
        with self.assertRaises(imapIO.IMAPError):
            self.email.as_string()

    def test_getitem(self):
        self.email['from']
        self.email['fromWhom']

    def test_setitem(self):
        self.email['from'] = ''
        self.email['fromWhom'] = ''


class IMAP4Dummy(imapIO._IMAPExtension):
    
    host = 'imap.mail.yahoo.com'
    port = ''
    user = ''
    error = Exception

    def xatom(self, a):
        pass


def test_build_message():
    imapIO.mimetypes.guess_type = lambda a: (None, None)
    imapIO.build_message(attachmentPaths=['MANIFEST.in'])
    imapIO.mimetypes.guess_type = lambda a: ('image/xxx', None)
    imapIO.build_message(attachmentPaths=['MANIFEST.in'])
    imapIO.mimetypes.guess_type = lambda a: ('audio/xxx', None)
    imapIO.build_message(attachmentPaths=['MANIFEST.in'])
    imapIO.mimetypes.guess_type = lambda a: ('xxx/xxx', None)
    imapIO.build_message(attachmentPaths=['MANIFEST.in'])


def test_normalize_nickname():
    assert imapIO.normalize_nickname('person.one@example.com') == 'Person One'
    assert imapIO.normalize_nickname('Mr. Person <person.one@example.com>') == 'Mr Person'


def test_utf_7_imap4():
    WORD = 'Спасибо'.decode('utf-8')
    assert WORD.encode(CODEC_NAME).decode(CODEC_NAME) == WORD
    assert 'one&'.encode(CODEC_NAME).decode(CODEC_NAME) == 'one&'
