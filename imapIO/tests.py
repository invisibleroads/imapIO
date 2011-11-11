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
            email.save(targetPath)
            partPacks = imapIO.extract_parts(targetPath)
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


class TestExceptions(unittest.TestCase):

    def setUp(self):
        self.serverDummy = IMAP4Dummy()

    def test_folders(self):
        with self.assertRaises(imapIO.IMAPError):
            self.serverDummy.cd = lambda: None
            self.serverDummy.list = lambda: ('xxx', [])
            self.serverDummy.folders
        self.serverDummy.list = lambda: ('OK', [''])
        self.serverDummy.folders
        self.serverDummy.list = lambda: ('OK', [('()', '"/"', 'xxx')])
        self.serverDummy.folders

    def test_cd(self):
        self.serverDummy.select = lambda: ('xxx', [])
        self.serverDummy.cd()

    def test_walk(self):
        self.serverDummy.capabilities = []
        with self.assertRaises(imapIO.IMAPError):
            self.serverDummy.walk(sortCriterion='ARRIVAL').next()
        self.serverDummy.cd = lambda a='': None
        self.serverDummy.list = lambda: ('OK', ['() "/" aaa', '() "/" bbb'])
        self.serverDummy.uid = lambda a, b, c, d: ('xxx', [])
        with self.assertRaises(StopIteration):
            self.serverDummy.walk('bbb').next()
        self.serverDummy.uid = lambda a, b, c, d=None: ('OK' if a.startswith('s') else 'xxx', ['1'])
        with self.assertRaises(StopIteration):
            self.serverDummy.walk('bbb').next()


class IMAP4Dummy(imapIO._IMAPExtension):
    
    user = ''
    host = ''
    port = ''
    error = Exception


def test_normalize_nickname():
    assert imapIO.normalize_nickname('person.one@example.com') == 'Person One'
    assert imapIO.normalize_nickname('Mr. Person <person.one@example.com>') == 'Mr Person'


def test_utf_7_imap4():
    WORD = 'Спасибо'.decode('utf-8')
    assert WORD.encode(CODEC_NAME).decode(CODEC_NAME) == WORD
