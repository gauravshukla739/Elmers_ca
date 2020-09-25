                  README for Morten's JavaScript Tree Menu
                      Version 2.3.0 dated 2001-04-30

 This software is released under the BSD License and it can be found
 in the file 'COPYING'.

 Changes to mtmcode.js from the previous version is found in 'ChangeLog'.

 The files 'code.html' and 'mtmcode.js' are the two main files needed to
 get anything working.  To find out more about installation, open the file
 'index.html' in either of the example sites in a web browser and follow
 the link named "Installation Guide".

 There used to be a file called 'commented-code.html' which had a commented
 version of 'code.html'.  It has been removed since I will be working on
 a detailed explanation on the web site under "How it works".

 The directory 'menu-images' contains all images that is used in the menu.
 Several icons that can be used is included.  The author would like more
 icons available, so if you have any you feel like sharing with others
 send them to Morten at the e-mail address <morten@treemenu.com> :)
   
 There's two example sites, one in each of the 'example-sites' directories.
 One uses a dynamically created frameset, the other uses a regular frameset.
 What this is and how it works is explained in the installation guide.
 Please make sure that you have updated all files before you publish them.
 The menu system goes to great lengths to make sure all users get content
 regardless of their setup.  Both example sites use 'code.html',
 'mtmcode.js', 'menu.html' _and_ 'index.html' to show content to the user.

Special note to upgrading users!:

 * Go through code.html and have a look at the options that have been changed
   or added.  Specifically from v2.2.6 to 2.3.0 they are:

   - MTMSubsGetPlus
   - MTMExtraCSS
   - MTMHeader
   - MTMFooter
   - MTMUseCookies
   - MTMCookieName
   - MTMCookieDays
   - MTMUseToolTips

 * The arguments to MTMItem() has changed.  The icon is now the 5th argument,
   and the new 4th argument is the tool tip text.

 * MTMakeSubmenu() now takes up to 4 arguments.  See the installation guide
   for details.

 * The tracking code has changed to work with Netscape 6 and Opera 5.

 Remember that the tree menu has it's own mailing list,
 tree_menu@yahoogroups.com.  Visit the web-site for information about how
 to join the list and receive updates and information regarding the menu 
 through e-mail.

 Bug-reports, comments, questions, feature suggestions, thank you notes
 and the like are all welcome by e-mailing Morten on <morten@treemenu.com>
