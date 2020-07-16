# Winium automated testing for Intact
--------
# features to add
-When the list gets shorter of things to do. Go back to the original ones and get the blanks and fill them in 

-Being able to run it completely outside of the dev environment. Can probably be done by having a console app that runs them using mstest or vstest

-When more than a couple tests fail, most likely it is going to be a problem with the program, maybe implement removing the first test that failed if array of test is larger than a certain amount. Could be done by creating a playlist and then deciding what playlist to run/uninclude test somehow.

-generate portfolios and that type of stuff

-On each screen go through and add interaction with each button and blank to see if it throws an error or messes something up

-Fix onFail so it will actually close out of the top layer window and if there is a save screen hit save and exit. verify that test resets are working on all screens!! <-- Hard

-Add annotations (Unimplemented Exeception) <-- Hard

-Add utilities support ( only some need help ) 

-In the fail file specify where the thing fails

-In final console app implement something to be able to clear out the specified folders/archive it.

-Figure out what else I need to add

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

-Search to determine if something is there

-Support for recognition

