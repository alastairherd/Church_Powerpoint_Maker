# Powerpoint_update
Project to auto-generate a Powerpoint from set variables

I will be expanding on this later, but here is a general outline of how the script works.

### Variables
This is a script that is used to generate the variables that will be used in the the other scripts. It is imported into:

### Functions
Here we have all our classes to be used later in the code, and at the bottom we have a another variable creation.
Because all these new variables are created using the above functions they live in this file rather than the previous variables file

### Slide_Making
Here we have the class for filling slides as well as functions to write to the slides. There is also a flag assignment for categorising what type of slide we are dealing with.

### Powerpoint
This is a nice and simple script that runs the previous scripts. It imports the template file and then writes into it, cycling through the variables created and writing them into the powerpoint. If the variable would not write properly to a slide then it just moves on.

Then it saves the newly named powerpoint at the end.

#### JSON Files
There are several JSON files that are used in this script:
1. psalms.json
   - This was manually created from the Free Church of Scotlands "Sing Psalms" zip file
2. components.json
   - This was manually created from a list of components that are used in our churches service
3. wsc.json
   - This comes from the "A Puritians Mind" website
   - Citations are in the file
