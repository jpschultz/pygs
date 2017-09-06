# PYGS - Python for Google Sheets

The purpose of this module is to allow a user to send a Pandas DataFrame to a google sheet as well as being able to read in a Google Sheet to a DataFrame.

## Getting Started

pip install git+https://github.com/jpschultz/pygs.git

### Prerequisites

You will need to setup a API Authentication Token with Google Sheets API and store the client_secret and a the json auth in specific folders. Here's how to do that:

1. Navigate to https://console.cloud.google.com/home/
2. At the top, to the right of where it say's 'Google Cloud Platform', and to the left of the search bar, you will see either a project you've created, or it will say 'Select a project'. If you would like to use a project you have already created, make sure it's selected and skip to step 6.
3. Click the + in the upper right to create a new project
4. Title your project what you want. Keep in mind, this project can be used for other Google Services/with other Google API's as well, it doesn't have to be just for this.
5. Select the radials and press 'Create'
6. With the project selected, on the left hand side, click 'APIs & services' then 'Library'
7. Locate the 'Sheets API' or you can search for it and then click on it.
8. At the top, click 'ENABLE'
9. While still in the 'APIs & services' menu, select 'Credentials', then click the blue 'Create credentials' button and select 'OAuth client ID'.
10. Select 'Other', give it a name of your choosing, and press the blue 'Create' button.
11. Press 'OK' to the OAuth client screen that pops up.
12. You should see a line item under 'OAuth 2.0 client IDs' where the name is the one you just created. All the way on the right, you will see a little download icon. Go ahead and download the client_secret json file and make sure to name it 'client_secret.json'
13. Create a 'clientsecrets' folder in your home directory and move your client_secret.json file to it. For instance, on a MAC, the folder would be located at ~/clientsecrets.

Congrats! The setup is (mostly) over. The rest will be automated when you import the module.

* Please note, if you are running this on a headless machine, you will need to run the import first on a desktop where you can access a browswer so you can finish the authentication. When you run 'import pygs' for the first time, it will open a browser window with to authorize the script. This only happens once and when the authorization is complete, it will create a folder called '.credentials' in your home directory where it will store the authentication json token. Copy the '.credentials' folder to the home directory of the headless machine and you should be able to import the module without any issues.


### Installing

Run the import and the first time, the import will create a '.credentials' folder in your home directory where it will store the result of the OAuth flow. You will be stored. You will see a browser window open asking you to authorize it and once that's complete, you will be able to use the module. This should only happen the first time you import the module.

```
import pygs
```


## Example uses

Creating a new spreadsheet from a dataframe:

```
pygs.create_spreadsheet_from_df(df, sheet_name='First Tab Name', document_name='Name of Newly Created Spreadsheet')
```


## Authors

* **JP Schultz** - *Initial work* - (https://github.com/jpschultz)



## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Google, thanks for making the quickstart!
