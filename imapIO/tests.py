# Import system modules
import unittest
import imaplib
import tempfile
import shutil
# Import custom modules
from scout.lib import imap
from scout import tests


class TestMailStoreIMAP(unittest.TestCase):

    temporaryFolderPaths = []

    def __init__(self, *args, **kwargs):
        'Load IMAP mailbox credentials for testing'
        # Call super constructor
        unittest.TestCase.__init__(self, *args, **kwargs)
        # Connect
        self.mailbox = imap.Store(*tests.credentials)

    def tearDown(self):
        'Delete temporary folders'
        for folderPath in self.temporaryFolderPaths:
            shutil.rmtree(folderPath, ignore_errors=True)

    def testReadPreservesFlags(self):
        """
        Make sure that a message that is marked unread is still marked unread
        even after reading the message.
        """
        # Get shortcuts
        server = self.mailbox.server
        messageCount = int(server.select()[1][0])
        messageIndices = range(1, messageCount + 1)
        # Define
        def countUnread():
            unreadCount = 0
            for messageIndex in messageIndices:
                messageFlags = imaplib.ParseFlags(server.fetch(messageIndex, 'FLAGS')[1][0])
                if r'\Seen' not in messageFlags:
                    unreadCount += 1
            return unreadCount
        # Mark each message as unread
        for messageIndex in messageIndices:
            server.store(messageIndex, '-FLAGS', r'(\Seen)')
        # Count the number of unread messages
        unreadCountBefore = countUnread()
        # Read each message
        for message in self.mailbox.read(includes=['inbox']):
            pass
        # Count the number of unread messages
        unreadCountAfter = countUnread()
        # Assert that the number of unread messages has not changed
        self.assertEqual(unreadCountBefore, unreadCountAfter)

    def testSavePreservesFlags(self):
        """
        Make sure that a message that is marked unread is still marked unread
        even after saving the message.
        """
        # Get a random message
        message = self.mailbox.read().next()
        # Mark the message as unread
        message.markUnread()
        # Save the message
        self.temporaryFolderPaths.append(tempfile.mkdtemp())
        message.save(self.temporaryFolderPaths[-1])
        # Assert that the message is still unread
        self.assertEqual(message.isUnread(), True)
