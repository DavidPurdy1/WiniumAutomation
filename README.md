# Winium automated testing for Intact
--------
# features to add
-When a test fails to rerun the test, taking a screenshot using the driver each time to show where it is going wrong, then taking those images and what elements are being found in the heirarchy and putting those on a document where they can be reviewed later.

-Being able to create the tests without code

-Being able to run it completely outside of the dev environment. 

-To be able to automate everything except the very extreme cases

-To add to an excel sheet what fails and what passes https://www.youtube.com/watch?v=yEa_Xo4yfHY

-Add an implicit wait class so the test works faster and better because then it will have a lot lower chance of something failing if it doesn't load in time.

-Driver closing on stopping a test or failing as well as actually failing the test when it can't find the element instead of just running for 60 seconds (Maybe either changing the inactivity time, but finding some way for when the test is stopped or failed for the driver to close) & when it passes

----------------------------------------------------------------------------------------------------------------

-Set up an InZone test on definitions that will recognize and prove that the definitions InZones are working. When it goes through InZone as long as the name is equal to one of the def names it will return true and pass the test]

-have it clear out the images from the failed image folder after it creates the documents so it doesn't add more images to the fail log. Have it rerun the test on failure, but figure out how to have it use the failLog method on each step.

-generate portfolios and that type of stuff

-On each screen go through and add interaction with each button and blank to see if it throws an error or messes something up

-Going through and testing every blank and scenario for the document attribution screen:: Can be reused 

# Features Added 

-Creation of Types and definitions and documents

-Loging in or connecting to a remote desktop (Set the path for the application and the user) 

-Adding Documents to InZone (set up the InZone definitions and documents in the right path for it to work)

-Adding Documents through batch review(Set up the definition and type for batch review)

-If failed can take screenshots then adds them to a word document

-Config file so it can be set up on any machine

-Default values for everything, but option to specify exactly what you want for most of the tests. Trying to balance ease and customization of the tests 


