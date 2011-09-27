'Tests for imapIO'
import os
import random
import tempfile
import unittest
import datetime
import ConfigParser
import logging; logging.basicConfig()

import imapIO


configuration = ConfigParser.ConfigParser()
if not configuration.read('.test.ini'):
    raise Exception('Please create a configuration file called .test.ini')
getX = lambda x: configuration.get('imap', x)
imap4 = getX('imap4').lower() == 'true'
imap4_ssl = getX('imap4_ssl').lower() == 'true'
host = getX('host')
port = int(getX('port'))
user = getX('user')
password = getX('password')


class ReplaceableDict(dict):
    
    def replace(self, **kwargs):
        return ReplaceableDict(self.items() + kwargs.items())


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
        # Get folders
        folders = self.server.folders
        # Test includes
        folder = random.choice(folders)
        tags = set(imapIO.parse_tags(folder))
        for email in self.server.walk(folder):
            self.assertEqual(tags.difference(email.tags), set())
        # Test excludes
        folder = random.choice(folders)
        tags = set(imapIO.parse_tags(folder))
        for email in self.server.walk(excludes=set(folders).difference(folder)):
            self.assertNotEqual(tags.intersection(email.tags), tags)
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
        baseCase = ReplaceableDict(
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
                'MANIFEST.in',
                'README.rst',
            ])
        # Clear previous cases
        for email in self.server.walk(includes=folder):
            if [email.fromWhom, email.toWhom] == [baseCase['fromWhom'], baseCase['toWhom']]:
                email.deleted = True
        self.server.expunge()
        # Run cases
        cases = [
            baseCase,
            baseCase.replace(bodyHTML=''),
            baseCase.replace(bodyText=''),
            baseCase.replace(attachmentPaths=None),
            baseCase.replace(attachmentPaths=None, bodyHTML=''),
            baseCase.replace(attachmentPaths=None, bodyText=''),
        ]
        self.temporaryPaths = []
        for caseIndex, case in enumerate(cases):
            subject = case['subject'] + str(caseIndex)
            # Revive
            self.server.revive(folder, imapIO.build_message(**case.replace(subject=subject)))
            # Make sure the revived email exists
            for email in self.server.walk(includes=folder):
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
        for email in self.server.walk(includes=folder):
            if [email.fromWhom, email.toWhom] == [baseCase['fromWhom'], baseCase['toWhom']]:
                self.server.revive(folder, email)
                break
        # Clear cases
        self.server.format_error('xxx', '')
        for email in self.server.walk(includes=folder):
            email.format_error('xxx', '')
            if [email.fromWhom, email.toWhom] == [baseCase['fromWhom'], baseCase['toWhom']]:
                email.deleted = True
        self.server.expunge()


@unittest.skipIf(not imap4, 'not configured')
class TestIMAP4(unittest.TestCase, Base):

    def setUp(self):
        if not hasattr(self, 'server'):
            self.server = imapIO.IMAP4.connect(host, port, user, password)


@unittest.skipIf(not imap4_ssl, 'not configured')
class TestIMAP4_SSL(unittest.TestCase, Base):

    def setUp(self):
        if not hasattr(self, 'server'):
            self.server = imapIO.connect(host, port, user, password)


def test_clean_nickname():
    assert imapIO.clean_nickname('person.one@example.com') == 'Person One'
    assert imapIO.clean_nickname('Mr. Person <person.one@example.com>') == 'Mr Person'
