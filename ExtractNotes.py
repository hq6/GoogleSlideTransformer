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

# Allow us to access page elements by id
idToPageElement = {}

def eliminateMatchingCriteria(e, filters):
    if type(e) is list:
        for x in e:
            eliminateMatchingCriteria(x, filters)
    if not type(e) is dict:
        return
    for key in e:
        if key in filters and filters[key](e[key]):
            del filters[key]
            continue
        eliminateMatchingCriteria(e[key], filters)

def getObjectIds(pageElements, processedFilters):
    # Output objectIds
    results = []
    # We can only find top-level objects for now.
    for element in pageElements:
      # Recursely check groups
      if "elementGroup" in element:
         results.extend(getObjectIds(element["elementGroup"]["children"], processedFilters))
         continue
      # Dig into the element until we find the field we desire.
      copyOfFilters = processedFilters.copy()
      eliminateMatchingCriteria(element, copyOfFilters)
      if not copyOfFilters: # All filters were matched
          print(element)
          results.append(element['objectId'])
    return results

def getService():
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    return build('slides', 'v1', http=creds.authorize(Http()))

def parseRanges(integerRanges):
    output = []
    for r in integerRanges.split(","):
      if '-' in r:
        start,end = r.split('-')
        for s in range(int(start), int(end) + 1):
            output.append(s)
      else:
        output.append(int(r))
    return output



def transformToRequest(transform, objectId):
    def default_to_regular(d):
        if isinstance(d, defaultdict):
            d = {k: default_to_regular(v) for k, v in d.items()}
        return d

    def recursive_defaultdict():
        return defaultdict(recursive_defaultdict)

    category, op = transform.split(":", 1)
    field, value = op.split("=")
    # Add more operation types here
    request = recursive_defaultdict()
    if category == "textStyle":
        request['updateTextStyle']['objectId'] = objectId
        request['updateTextStyle']['textRange']['type'] = 'ALL'
        request['updateTextStyle']['fields'] = field
        # Add more types of text style changes here
        if field == 'fontSize':
            request['updateTextStyle']['style']['fontSize'] = {'magnitude' : value, 'unit' : 'PT'}
        elif field == 'fontColor':
            request['updateTextStyle']['fields'] = 'foregroundColor'
            request['updateTextStyle']['style']['foregroundColor']['opaqueColor']['themeColor'] = value
    elif category == "shapeProperties":
        request['updateShapeProperties']['objectId'] = objectId
        request['updateShapeProperties']['fields'] = field

        # Generate request from field and value
        fieldParts = field.split(".")
        cur = request['updateShapeProperties']['shapeProperties']
        for x in fieldParts[:-1]:
            cur = cur[x]

        # Assign the value.
        if value.startswith("&"):
            # Grab the value from the target object
            value = idToPageElement[value[1:]]
            # print("FOO " + str(value) + " BAR" )
            value = value['shape']['shapeProperties']
            for x in fieldParts:
                value = value[x]
        # print("FOO " + str(value) + " BAR" )
        cur[fieldParts[-1]] = value
    elif category == 'transformPageElement':
        # Currently support only reference elmeents
        if not value.startswith("&"):
            print("Transforming page elements only supports reference objects!", file=sys.stderr)
        refId = value[1:]
        request['updatePageElementTransform']['objectId'] = objectId
        request['updatePageElementTransform']['applyMode'] = "ABSOLUTE"
        request['updatePageElementTransform']['transform'] = idToPageElement[refId]['transform']

    return default_to_regular(request)



def generateIdToPageElement(elements):
    output = {}
    for element in elements:
      if "objectId" in element:
        output[element['objectId']] = element
      if type(element) == dict:
        for key in element:
          output.update(generateIdToPageElement(element[key]))
      if type(element) == list:
        for o in element:
          output.update(generateIdToPageElement(o))
    return output

def tryParse(element):
    # print(element['objectId'])
    # print(element['shape']['text']['textElements'][1]['textRun']['content'])
    textElements = []
    try:
      textElements = element['shape']['text']['textElements']
    except:
      print("Failed to get element['shape']['text']['textElements']")

    # print(textElements)
    for text in textElements:
      try:
        print(text['textRun']['content'].encode('utf-8'))
      except:
        pass

doc = r"""
Usage: ./ExtractNotes.py <presentation_id>

    -h,--help                    show this
"""
def main():
    options = docopt(doc)
    print(options)

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

    # Build an index of objectIds to PageElements
    global idToPageElement
    idToPageElement = generateIdToPageElement([element for slide in slides for element in slide.get('pageElements')])
    print('The presentation contains {} slides.'.format(len(slides)))

    # Extract notes.
    for i, slide in enumerate(slides):
        # Find all elements matching the criteria
        print("Slide {0}".format(i+1))
        notesPage = slide["slideProperties"]["notesPage"]
        notesId = notesPage["notesProperties"]["speakerNotesObjectId"]
        # print(notesId)
        for element in notesPage['pageElements']:
            # print(element['objectId'])
            if element['objectId'] == notesId:
              # print(element)
              # print(element['shape']['text']['textElements'][1]['textRun']['content'])
              tryParse(element)

        print("=" * 80)
        # break

if __name__ == '__main__':
    main()
