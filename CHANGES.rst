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
