# Winium automated testing for Intact
--------
# features to add
-When a test fails to rerun the test, taking a screenshot using the driver each time to show where it is going wrong, then taking those images and what elements are being found in the heirarchy and putting those on a document where they can be reviewed later.

-Being able to create the tests without code

-Being able to run it completely outside of the dev environment. 

-To be able to automate everything except the very extreme cases

-Removed Microsoft.Office.Interop

-When more than a couple tests fail, most likely it is going to be a problem with the program, maybe implement removing the first test that failed if array of test is larger than a certain amount. Could be done by creating a playlist and then deciding what playlist to run/uninclude test somehow.

-have it clear out the images from the failed image folder after it creates the documents so it doesn't add more images to the fail log. Have it rerun the test on failure, but figure out how to have it use the failLog method on each step.

-generate portfolios and that type of stuff

-On each screen go through and add interaction with each button and blank to see if it throws an error or messes something up
 
-verify that test resets are working on all screens!! <-- Hard


-Add to search 

-Fix onFail so it will actually close out of the top layer window and if there is a save screen hit save and exit

-Add annotations (Unimplemented Exeception) <-- Hard


-Add utilities

-Fix recognize so it will work completely <-- Hard

-In the fail file specify where the thing fails

# Features Added 

-Creation of Types and definitions and documents

-Loging in or connecting to a remote desktop (Set the path for the application and the user) 

-Adding Documents to InZone (set up the InZone definitions and documents in the right path for it to work)

-Adding Documents through batch review(Set up the definition and type for batch review)

-If failed can take screenshots then adds them to a word document

-Config file so it can be set up on any machine

-Default values for everything, but option to specify exactly what you want for most of the tests. Trying to balance ease and customization of the tests 

-Inzone recoginizes correct definitions when coming in. 

-Fail file to tell you which tests failed and passed

-Logout and log back into intact testing

