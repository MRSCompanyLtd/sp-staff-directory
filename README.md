# sp-staff-directory

## Summary

SharePoint staff directory webpart built using SPFx and React.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.13-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| src/webparts/staffDirectory | [MRS Company Ltd.](https://github.com/MRSCompanyLtd) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | September 28, 2022 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:

```bash
yarn # install dependencies
```

```bash
gulp serve # run dev environment
```

To build and deploy:

```bash
gulp build # build solution
gulp bundle --ship # bundle solution for production
gulp package-solution --ship # create sppkg file for deploy
```

Then upload the sppkg file from /sharepoint folder into app catalog site and allow Graph API in the admin portal.

## Features

This webpart offers the following functionality

- Get all users in the tenant
- Search user by first name, last name, department, or job title.
- Search by letter.
- Filter by department.  

## Contact

Contact MRS Company on our [website](https://mrscompany.com), [Twitter](https://twitter.com/_MRSCompany), or [GitHub](https://github.com/MRSCompanyLtd).
