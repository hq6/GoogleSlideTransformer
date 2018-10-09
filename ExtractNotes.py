#!/usr/bin/python

# Compatibility with Python3
from __future__ import print_function

# Imports for Google APIs
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# Parsing options
from docopt import docopt

from sys import exit, argv

# Nested dictionaries
from collections import defaultdict

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/presentations.readonly'

def getService():
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    return build('slides', 'v1', http=creds.authorize(Http()))

def tryParse(element):
    textElements = []
    try:
      textElements = element['shape']['text']['textElements']
    except:
      pass

    # print(textElements)
    for text in textElements:
      try:
        print(text['textRun']['content'].encode('utf-8').strip())
      except:
        pass

doc = r"""
Usage: ./ExtractNotes.py <presentation_id>

    -h,--help                    show this
"""
def main():
    options = docopt(doc)

    # Clear argv so that the oauth service does not freak out
    del argv[1:]
    service = getService()

    # Call the Slides API
    try:
        presentation = service.presentations().get(
            presentationId=options['<presentation_id>']).execute()
        slides = presentation.get('slides')
    except:
        print("Failed to get slides, likely invalid presentation id passed")
        exit(1)

    # Extract notes.
    for i, slide in enumerate(slides):
        # Find all elements matching the criteria
        print("=" * 80)
        print("Slide {0}".format(i+1))
        notesPage = slide["slideProperties"]["notesPage"]
        notesId = notesPage["notesProperties"]["speakerNotesObjectId"]
        for element in notesPage['pageElements']:
            if element['objectId'] == notesId:
              tryParse(element)

if __name__ == '__main__':
    main()
