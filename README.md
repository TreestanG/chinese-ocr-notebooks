# 📓 Chinese Notebook Vocab
A little script I made in order to automate some a weekly assignment at school.
It uses the google slides api to interact and update tables. 

## Usage
1. Folow [these](https://developers.google.com/slides/api/quickstart/python) instructions and download `credentials.json` into this folder.
2. Set the values in `config.json` to match your likings.
3. Open `index.py` and update the vocabulary text at the top. (Leave the space following `"""` empty)
4. Run the code!

Keep in mind that the slides containing the tables must be in this format:
`PageElement`, `Table`, `Table`. If your vocab list contains more entries than the allocated table, split the vocab list into 2 slides and repeat on a new slide. 

To switch accounts, just delete the `token.json` file (which will be created if you have not logged in before).