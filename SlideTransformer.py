#!/usr/bin/python

# Compatibility with Python3
from __future__ import print_function

# Imports for Google APIs
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# Parsing options
from docopt import docopt

from sys import exit

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/presentations'

# The ID of a sample presentation.
PRESENTATION_ID = '1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI'

def eliminateMatchingCriteria(e, filters):
    if not type(e) is dict:
        return
    for key in e:
        if key in filters and e[key] == filters[key]:
            del filters[key]
            continue
        eliminateMatchingCriteria(e[key], filters)

def getObjectIds(pageElements, rawFilters):
    # First parse rawFilters. They are ANDed together logically.
    processed_filters = {}
    for x in rawFilters:
        key, value = x.split("=")
        processed_filters[key] = value

    # Output objectIds
    results = []
    # We can only find top-level objects for now.
    for element in pageElements:
      # Dig into the element until we find the field we desire.
      copyOfFilters = processed_filters.copy()
      eliminateMatchingCriteria(element, copyOfFilters)
      if not copyOfFilters: # All filters were matched
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


doc = r"""
Usage: ./SlideTransformer.py  [-s <slide_ranges>] [-f <filter>]... [-t <transform>]... <presentation_id>

    -h,--help                    show this
    -f,--filters <filter>        filters of form key = value; Future implementations may generalize this
    -t,--transform <transform>   Transform using the syntax required by Google Slides API, without objectId.
    -s,--slides <slide_ranges>   Slide ranges to look for perform transformations over.
"""
def main():
    options = docopt(doc)
    print(options)
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

    # Check for a slide filter
    slideFilter = None
    if options['--slides']:
        slideFilter = parseRanges(options['--slides'])

    requests = []
    print('The presentation contains {} slides:'.format(len(slides)))
    for i, slide in enumerate(slides):
        # Slides are numbered starting at 1
        if slideFilter and (i + 1) not in slideFilter:
            continue

	# Find all elements matching the criteria
	objectIds = getObjectIds(slide.get('pageElements'), options['--filters'])
	print(objectIds)

        #if pageNumId:
        #    request = {}
        #    request['updateTextStyle'] = {}
        #    request['updateTextStyle']['objectId'] = pageNumId
        #    request['updateTextStyle']['textRange'] = { 'type': 'ALL' }
        #    request['updateTextStyle']['style'] = { 'fontSize': {'magnitude' : 14, 'unit' : 'PT'}}
        #    request['updateTextStyle']['fields'] = 'fontSize'
        #    requests.append(request) 

    # Send request
    #body = { 'requests': requests }
    #response = service.presentations().batchUpdate(
    #            presentationId=PRESENTATION_ID, body=body).execute()



if __name__ == '__main__':
    main()
