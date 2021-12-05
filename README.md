# Wilco-Scripting
## Introduction
Public version of my repository used at work to showcase some of my first bits of PowerShell scripting and development.  
  
The results of creating this module is what led to my success in my role as a Service Desk Technician. It also further sparked my interest for the use of source control on Github, and software development in General. 
## Modules Used
[ImportExcel](https://github.com/dfinke/ImportExcel) - Module used to read, write, and manipulate Excel Files.  
[ActiveDirectory](https://docs.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2019-ps) - Module used to interface with Active Directory.  
[MSOnline](https://www.powershellgallery.com/packages/MSOnline/1.1.183.66) - Module used to interface with multiple Office365/Microsoft365 components.  
[AzureAD](https://docs.microsoft.com/en-us/powershell/module/azuread/?view=azureadps-2.0) - Module used to interface with Azure Active Directory. At the time of development, Wilco was in a hybrid state, requiring the use of both Azure AD and on-prem AD.  
[ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/2.0.5) - Module used to assign permissions to all types of Office365/Microsft365 mailboxes
## Outcomes  
By creating these scripts, I was able to achieve almost complete hands-off automation in repetitive tasks I was doing as an entry level Service Desk Technician many times per day. This involved:  
### Common Operations Done Daily
- Creating Users in Active Directory
- Assigning Distribution and Security Groups in on-prem AD and Azure AD
- Assigning Office365 Licenses
- Placing users in the appropriate OU
- Setting All AD Attributes
- Generating an email to send to appropriate team members after user creation
- Changing user permissions if they had a change in job responsibilities
- Separating users from the company
### Less Common Operations Also Automated
- Changing "Reports To" field in AD if a Manager Separated
- Generating Reports of Office Licensing Information combined with other AD Attributes
- Generating Reports of other specific AD Attributes
- Auditing AD Monthly to ensure AD Contents Match HR's Database  
- Auditing all user permissions

It also led to enhancing some of the data integrity of users in Active Directory.

## The Process, Before and After
### Before
Prior to these scripts, the process to create a user was to enter 4 pieces of information to a PowerShell Script. The script would then assign location specific permissions, assign the name attributes, email attributes and the job title. Other aspects of the user setup would be assigned manually. Once the user was created, an email template would need to be edited with the appropriate information to send.
### After
The scripts accept the full email from HR requesting a new user, separation, or job role change, the scripts will parse the information for the required variables and send it down the pipeline to the appropriate script which will then:
- Reads in data from an Excel File defining all job roles 
- Validate all 4 pieces of information to ensure they match a defined position
- Create the user and assign all attributes like before
- Setup the enitrety of the user automatically
- Output a txt file that contains an email and the email recipients. The email contains all required info.

The whole process to create a user typically took about 5 minutes and was error prone. With this implementation, The whole process is extremely accurate and is typically completed in under 60 seconds.

## Conclusion
I was very impressed with the implementation Wilco had when I arrived, but once I learned some of the capabilities of PowerShell, I was excited to try to improve upon what had already been built!  

Some of the information I will be unable to include because it would reveal business specific information. This is mainly a repository to showcase my first "large" project that I completed almost entirely by myself. Building these scripts, along with taking an Open Source Software Class inspired me to make my first OSS contribution to ImportExcel, where I contributed documentation that would have helped me when I was first learning the module. The documentation I contributed can be seen [here](https://github.com/dfinke/ImportExcel/tree/master/FAQ)😊
