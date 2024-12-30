# Google-Docs-Participation statistics

Make use of Google docs api and retrieve revisions and historical growth of a google document.

Create an analytics report of time, participants, who contributed the most (in terms of words, lines, ..), who commented the most, who resolved comments the most,. . . . .


# Install

1. Need to set up Google Cloud credentials:

- Go to the Google Cloud Console
- Create a new project or select an existing one
- Enable the Google Docs API and Google Drive API
- Create OAuth 2.0 credentials
- Download the credentials and save them as credentials.json in the same directory as the script

2. Create a virtual environment:
```
python -m venv .venv
source .vevn/bin/activate
pip install -r requirements.txt
```
# Run

Use a google doc that you have edit access: for example of this doc: https://docs.google.com/document/d/1P6ZEX1kpC85Mw_wBPk8ccGlHg1PlqmV2Yju4ZOLkQVo/edit?usp=sharing . Obviously, this script makes more sense when you have lots of edits, users, and activity history. Then just run the script with:

```
python gdocs_analyzer.py 1P6ZEX1kpC85Mw_wBPk8ccGlHg1PlqmV2Yju4ZOLkQVo
```

# Outputs

You will get (printed in the command line)
- Document Information (Title, Created time, Modified Time, Owner, Last modifier)
- Word Count Statistics over each user
- Word Count history over each revision
- User contributions (num of revisions)
- Estimated word contributions
- Comment activity
- Summary statistics (total edits, comments, replies, resolved comments, current word count, avg words per edit)

And also historical plots of the above in document_metrics.html

