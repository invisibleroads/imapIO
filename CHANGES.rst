0.9.4
-----
- Modified walk() to accept a generic function to filter folders
- Modified Email.__init__() to apply _decode() to both parts of an email address
- Removed clean_tag(), parse_tags, format_tags()
- Added utf-7-imap4 codec to parse folder names
- Increased test coverage to 100%


0.9.3
-----
- Fixed revive() to handle messages that lack a date
- Modified Email so that an email from walk() can be sent directly to revive()
- Modified Email so that we can directly access its parent folder
- Modified flags.setter so that it does not try to set flag "\Recent"


0.9.2
-----
- Reverted to set() for versions of Python < 2.7 that lack set literal syntax
- Removed keyword arguments from decode() to support versions of Python < 2.7
- Fixed tests for servers like Lotus Domino that do not update search indices


0.9.1
-----
- Changed walk() to use UID directly
- Added support for sortCriterion using UID SORT
- Improved test coverage to 80%


0.9.0
-----
- Extracted code from imap-search-scout
- Made API more user-friendly
- Improved test coverage to 79%
