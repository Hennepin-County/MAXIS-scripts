BlueZone Scripts | DHS-MAXIS-Scripts
===

Introduction
---

Welcome to the GitHub repository and project site for the MAXIS BlueZone Scripts! This project aims to automate repetetive, error prone tasks using simple extensions to the BlueZone Mainframe Display system. These scripts do not contain any confidential data, nor do they contain information about how to log in to our various state systems. 

If you have questions about BlueZone Scripts and work in a Minnesota human services agency, please ask a supervisor about getting started.

GitHub workflow and organization
---

GitHub can be somewhat complex and daunting for beginners. For our organization, scripts are divided into two "branches":

* **MASTER**: the working directory for scriptwriters and select power users in scriptwriting agencies.
* **RELEASE**: the branch for most eligibility workers statewide.

Changes proposed here (assuming they "pass muster" with any policy/procedural folks involved at DHS) will first be merged into "master", then into "release" after **at least** a week of testing. **Scriptwriters (and a few select "power users" in each agency) are expected to work off of the master branch, and _contribute feedback_ throughout the month**. The recommended procedure is to give all master branch users access to both master and release versions of the scripts (using separate installations and "ZMD" configuration profiles). Agencies may stray from this procedure, but it is not recommended (as the master branch is the statewide "test" branch).

The newest/upcoming draft of release notes (upcoming changes that have already been built) is [located here](https://gist.github.com/theVKC/16fea8523efbb3df1917). Scriptwriters and power users are encouraged to "star" this document to get updates on the newest changes as soon as they are available for testing.

Issue list
---

We have an [issue list](https://github.com/MN-Script-Team/DHS-MAXIS-Scripts/issues) maintained on GitHub. Both scriptwriters and non-scriptwriters should feel free to create/report issues on the issue list (doing so requires a GitHub account). 

Scriptwriters are encouraged to tackle any issues on the issue list, so long as they meet the following conditions: 
* The scriptwriter has time in the near future to complete and test the issue (note that many issues have a "milestone", which may have an associated due date).
* The scriptwriter adds a comment to the issue saying they'll take it.
* The scriptwriter builds (or modifies) instructions on SIR after completing the work.
* For new scripts, the scriptwriter tests the new script on multiple cases/scenarios before submitting (ideally, for a week or so on active cases in their agency).
 
**ABSOLUTELY NO CLIENT DATA SHOULD EVER BE SHARED ON GITHUB.** In addition, please refrain from posting entire screenshots of system screens on GitHub issues. If case numbers or screenshots are needed, please share them via secure email (see your agency for your local process).

### Issue guidelines/best practices
* Search existing issues before submitting a new one. Duplicates are annoying and add unneccessary work for administrators (as well as duplicate email notifications). It may also be wise to search through closed issues (by selecting "closed" in the top of the issue list).
* Issue title should be short (under 75 characters, or about the size of a case note header). This goes in the subject line for emails, so keep it clean.
* For existing scripts, please indicate the script category/name at the beginning of the issue (ex. "NOTES - CAF: needs longer space for 'other notes'"). This is helpful for organization.
* If there are multiple issues with an existing script, create separate issues for each. This is easier both for release notes tracking and for recipients of GitHub update emails.
* Don't upload screenshots of code, as it does not meet accessibility standards (and can't easily be copy/pasted). If you want to discuss code snippets, copy/paste them and surround them in blocks using GitHub markdown's default format (3 backticks: ```).
* If you have a question, it should only be posted if you believe a change to a script is necessary or wise. If it's a general scripts question, it is better addressed on the SIR discussion forum or via email.

### Critical issues
Sometimes a bug or enhancement needs to be prioritized over other issues. We can mark those issues as "critical", which gets our attention. Here are the two situations in which a bug or enhancement is considered "critical":
* A script **in the release branch** has an inhibiting edit which is completely impassible.
* A script **in the release branch** is doing an action which has been (or will be) considered "incorrect policy".
 
_Please note_: script issues that only occur in the master branch are _not_ considered critical, as _the master branch is considered a test ground_. For this reason, **it is recommended that all master branch users also have access to the release branch**.

### Script freezes
Script freezes are needed for making sure each new script, bug fix, and enhancement is properly tested. Generally speaking, enhancement/new script freezes are in effect in the following instances:
1. During the third week of each month (which corresponds roughly to the week before a script release).
2. When the issue list is over 40 issues (in this case, completing existing issues on the list is acceptable, so long as the first condition is respected).
3. For a few days prior to a major, project-wide update (such as a Functions Library update).

When there are over 40 issues, no new scripts or enhancements will be allowed on the GitHub issue list, unless they are critical from a policy standpoint (bug fixes are always welcome). Administrators may institute script freezes at other times dependent on need, and in these cases an email will be sent to scriptwriters.

Other
---

This is a casual (and, dare I say, fun) group, but we're doing a pretty epic project, with the goal of saving thousands of hours in work for all of Minnesota. So please be ready with a working knowledge of VBScript, an understanding of Minnesota's assistance programs, and a sense of humor, and I'm sure this will go well.

If you don't have a great working knowledge of VBScript, you can find out more by checking out the Wiki on this repository!

Oh, and don't be surprised if I'm pushing commits late in the evening. Frankly, I'm pretty obsessed with this project. ;)

Welcome to it!

Veronica Kahl Cary (VKC), Project Manager and SNAP Data Analyst, Minnesota DHS
