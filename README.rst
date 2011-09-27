imapIO
======
Here are some convenience classes and methods for processing IMAP mailboxes.  Since the classes are derived from the ``imaplib`` classes, all methods available in the ``imaplib`` classes are directly usable.


Installation
------------
::

    easy_install -U imapIO


Usage
-----
::

    # Connect to IMAP server
    import imapIO
    server = imapIO.connect(host, port, user, password)

    # Select folder
    import random
    messageCount = server.cd(random.choice(server.folders))

    # Walk messages in inbox sorted by arrival time
    for email in server.walk(includes='inbox', sortCriterion='ARRIVAL'):
        # Show information
        print
        print 'Date: %s' % email.whenUTC
        print 'Subject: %s' % email.subject.encode('utf-8')
        print 'From: %s' % email.fromWhom.encode('utf-8')
        print 'From (nickname): %s' % imapIO.clean_nickname(email.fromWhom)
        print 'To: %s' % email.toWhom.encode('utf-8')
        print 'CC: %s' % email.ccWhom.encode('utf-8')
        print 'BCC: %s' % email.bccWhom.encode('utf-8')
        # Set flags
        email.seen = False
        email.deleted = False

    # Walk messages satisfying search criterion
    emailCriterion = 'BEFORE 23-JAN-2005'
    emailGenerator = server.walk(excludes=['public', 'trash'], searchCriterion=emailCriterion)
    for emailIndex, email in enumerate(emailGenerator):
        # Show flags
        print
        print email.flags
        # Save email in compressed format on hard drive
        emailPath = '%s.gz' % emailIndex
        partPacks = email.save(emailPath)
        # Extract attachments from email on hard drive
        for partIndex, filename, contentType, payload in imapIO.extract_parts(emailPath):
            print len(payload), filename.encode('utf-8')

    # Create a message in the inbox
    import datetime
    server.revive('inbox', imapIO.build_message(
        whenUTC=datetime.datetime(2005, 1, 23, 1, 0),
        subject='Subject',
        fromWhom='from@example.com',
        toWhom='to@example.com',
        ccWhom='cc@example.com',
        bccWhom='bcc@example.com',
        bodyText=u'text',
        bodyHTML=u'<html>text</html>',
        attachmentPaths=[
            'CHANGES.rst',
            'README.rst',
        ]))
    email = server.walk('inbox', searchCriterion='FROM from@example.com TO to@example.com').next()
    email.deleted = True
    server.expunge()

    # Duplicate a message from one server to another
    server1 = imapIO.connect(host1, port1, user1, password1)
    server2 = imapIO.connect(host2, port2, user2, password2)
    server2.revive('inbox', server1.walk().next())
