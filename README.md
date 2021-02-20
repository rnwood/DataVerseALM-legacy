# DataVerseALM
 This is a repo template for a lightweight (no CI server) ALM solution for a PowerApps/DataVerse (including D365 Online) solution and associated system/configuration data.
 
 ## Features

- Git source control of a PowerApps solution using "Solution Packager" tool from MS
- source control of important configuration data
- better support for isolation of developers and branching and merging than other solutions
- repeatable "single artifact" versioning, release and deployment - release once, keep the generated files for version X.Y.Z and deploy 0..N times to different environments.
- since it's just PowerShell scripts, fully extendable.
- downloads the needed dependencies automatically, so easy to share.

See Microsoft documentation on the best practices this is designed to support:

- [Application lifecycle management](https://docs.microsoft.com/en-us/power-platform/admin/wp-application-lifecycle-management)
- [Use source control with solution files](https://docs.microsoft.com/en-us/dynamics365/customerengagement/on-premises/developer/use-source-control-solution-files)

 ## To use this template:

 - Copy the contents of this repo into a new Git workspace
 - Edit this file and remove this header. 
 - Add any custom first time setup instructions (like selecting 'yes' to the 'enable D365 apps' option if your solution needs is)
 - Edit `settings.ps1` and set the `$solutionname` and list of dependencies (put files in `dependencies`)
 - Edit `data\config.json` to define the set of data records to be controlled.
 - Follow the instructions below to understand how to export the initial version of your solution and data.

## FAQ

> **Isn't it easier/better to use an Azure Devops pipeline with Microsoft Power Platform Build Tools?**

Like most things there are pros and cons and your experience may vary depending on your team and requirements:

| Area | Azure DevOps pipeline w/PPBT | DataVerseALM |
|-|-|-|
| Maker/developer skills needed | Awareness of AzDo pipeline and support from administrator when things go wrong | Source Control, Running the scripts |
| Administrator skills needed | AzDo Pipelines, Source Control | Editing PowerShell scripts |
| Isolation between developers | Hard to setup/use with more than one dev environment ([see here for an approach](https://www.itaintboring.com/dynamics-crm/cicd-for-powerplatform-round-3/)) | Devs can work in own environment by default  |
| Visibility/control of changes before commit | None by default as changes are commited without review (although can be achieved with a PR process) | Changes available for review before committing |
| Usefullness of the source control history | Difficult due to instability of the generated XML outputs | Improved because of 'canonicalisation' process |
| Ability to branch and merge | Possible but difficult due to instability of the generated XML outputs | Improved because of 'canonicalisation' process |

In summary the AzDO approach is more suited to simple 'maker' scenarios (but needs someone with more experience to set it up and support them) and DataVerseALM approach might work better for pro-development teams (but needs some investment in source control skills to use).

---
*** End of template header. Remove everything above the --- below when you use this template ***
---

# How To Develop, Release and Deploy the Solution in This Repository #

The files in this repo include a PowerApps Solution and supporting configuration data which have been expanded into individual files using a tool.

To work on this, you need to import them into a PowerApps DataVerse environment. You can then make changes to the solution and/or data before exporting the changes back into the files which are then checked into source control.
You must work on this in a clean environment which is your own and not shared by other people.

## How to import the solution and data so you can view and edit it ##
- Get the latest version of all the files from source control into your repository. e.g. `git pull`
- Create an empty PowerApps environment with a DataVerse database
- Open an PowerShell prompt and execute the following command:

  `& pathto\import.ps1 -verbose`

  Where `pathto` is the path to your workspace folder for this repo. 
- The first time you run the script, it will prompt you to connect to a DataVerse environment. Connect to the organisation you've setup for development. Subsequent invocations in the same PowerShell session will remember the connection which was selected. 
- The script will now:
  - Deploy the correct version of Core dependency (not if you're working on Core itself obviously)
  - Import the unmanaged DataVerse solution from the `solution` subfolder. (This uses the Solution Packager tool provided by the PowerApps SDK to turn these files into a normal `.zip` solution file.)
  - Import the configuration data from the `data` subfolder.

## Editing the DataVerse solution

After importing the latest version you can edit the PowerApps solution using the normal facilities in the PowerApps Maker UI.

Once completed, don't forget to add any new components to the solution.

## Editing the configuration data

After importing the latest version, you can edit the configuration data in your PowerApps environment by using the normal forms etc.

## Exporting the PowerApps solution and data after editing

- Make sure you have the same version of the PowerApps solution and data files in your workspace that was used to import into your environment. This is important because when you export, it will overwrite and lose any changes which may exist there (e.g from other people).

- Open an PowerShell prompt and execute the following command:

  `& pathto\export.ps1 -verbose`

  Where `pathto` is the path to your workspace folder for this repo. 

- The script will prompt you to connect to the PowerApps environment if you hae not already connected in the PowerShell session.

- The script will then export the solution and data changes you have made into the files in ``solution`` and ``data`` subfolders. You should check the changes in these files to check they're as expected and complete.

- You can then commit the changes you've made into source control in whatever way you're comfortable with (so using a Git GUI or command line). 

## Creating a Release ##

Note that 'release' in this context means a package of files with an assigned version number.
Following this process does not apply the changes to any environment. That process is called 'deployment' - see following section. The release should be done once and then the resulting package stored used to deploy to each environment that needs that version.

- Ensure that the solution version number has been set to match the desired version number for the release.

- Ensure that your development environment has the latest version imported and does not have any additional untracked changes. (The release process has to use this to generate a managed solution).

- Make sure you have no untracked changes in your git workspace.

- Open an PowerShell prompt and execute the following command:

  `& pathto\release.ps1 -verbose`

  Where `pathto` is the path to your workspace folder for this repo. 

- Release will be generated at `..\.release` (not committed). Copy these to an shared archive and keep safe. You should use the same exact files to deploy to each environment to help make sure that what you test in your test env is exactly what gets deployed in production.

## Deploying a release ##

- Download files for release

- Open an PowerShell prompt and execute the following command:

  `& pathto\deploy.ps1 -verbose`

  Where `pathto` is the path to the files downloaded.

## Doing release followed by deploy to one env immediately ##

- As a convenience to speed up iteration, you can do a release followed by immediate deploy to a given environment by running the following:
`& pathto\release.ps1 -deployurl https://myenv.crm11.dynamics.com -commitcomment "#123 Change something important"`
  Where `deployurl` is the URL of the environment to deploy to. and `commitcomment` is the git commit comment for the changes.

  This is equivalent to running `./release.ps1` followed by `./deploy.ps1` on the generated release.

  Note that this should only be used for the first environment that needs to be deployed to (e.g. the test env). Re-use the release you already created when deploying to subsequent environments (e.g. the live env) so that what was tested is exactly what gets deployed.
