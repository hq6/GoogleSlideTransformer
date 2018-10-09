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

# Find slide boundaries
import re

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/presentations'

def getService():
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    return build('slides', 'v1', http=creds.authorize(Http()))

def parseNotes(notesFile):
    result = {}
    currentLines = []
    possibleNewSlide = False
    previousSlideNum = None
    with open(notesFile) as notes:
      for line in notes:
        line = line.rstrip()
        currentLines.append(line)
        if line.startswith("=" * 80):
          possibleNewSlide = True
          continue
        if possibleNewSlide:
          # See if a new slide matches
          match = re.match(r"Slide ([1-9]\d*)", line)
          if match:
            slideNum = int(match.group(1))
            if previousSlideNum:
              result[previousSlideNum] = "\n".join(currentLines[:-2])
            previousSlideNum = slideNum
            currentLines = []

          possibleNewSlide = False
    # The last slide should be added
    if previousSlideNum and currentLines:
        result[previousSlideNum] = "\n".join(currentLines)
    return result



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
Usage: ./InsertNotes.py <presentation_id> <notes>

    -h,--help           show this
    <presentation_id>   the presentation to update
    <notes>             A text file with records beginning with 80 '=' characters, and Slide N.
"""
def main():
    options = docopt(doc)

    # Parse the notes
    notes = parseNotes(options['<notes>'])

    # Test what was read
    for x in range(1, max(notes.keys()) + 1):
        if x not in  notes: continue
        print("=" * 80)
        print("Slide {0}".format(x))
        print(notes[x])
    return

    # Clear argv so that the oauth service does not freak out
    del argv[1:]
    service = getService()

    # Call the Slides API
    try:
        print("Processing presentation {0}".format(options['<presentation_id>']))
        presentation = service.presentations().get(
            presentationId=options['<presentation_id>']).execute()
        slides = presentation.get('slides')
    except:
        print("Failed to get slides, likely invalid presentation id passed")
        exit(1)


    print('The presentation contains {} slides.'.format(len(slides)))

    # TODO: Build the requests

    # Extract notes.
    for i, slide in enumerate(slides):
        # Find all elements matching the criteria
        print("Slide {0}".format(i+1))
        notesPage = slide["slideProperties"]["notesPage"]
        notesId = notesPage["notesProperties"]["speakerNotesObjectId"]
        for element in notesPage['pageElements']:
            if element['objectId'] == notesId:
              tryParse(element)

        print("=" * 80)

    # TODO: Send the requests

if __name__ == '__main__':
    main()
