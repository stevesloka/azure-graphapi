# azure-graphapi change log

## Version 0.0.8

- Changed strformat dependency from version 0.0.5 to 0.0.7.
- Updated comments.
- Added .DS_Store to .gitignore.

## Version 0.0.7

- Changed the main property in package.json to the correct case.

## Version 0.0.6

- Status code was not attached to error object from httpRequest.

## Version 0.0.5

- Access token was not being set correctly on retry.

## Version 0.0.4

- Fixed bug in the _getObjects method.
- Removed the generic request method.
- Fixed bug when appending version to existing query string.
 
## Version 0.0.3

- Allow string formatting using placeholders (require strformat).
- Fixed bug in getObjects where get was not called with a context.
- Remove log object reference.
- Fixed HTTP status code lookup (require https).

## Version 0.0.2

- Corrected typographical errors in the README.md file.
- Corrected the repository in the package.json file.
- Corrected comments in the GraphAPI.js file.
- The callback data is the response.value and not the response itself.

## Version 0.0.1

- Initial commit.
