#!/usr/bin/python

from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/presentations.readonly'

# The ID of a sample presentation.
PRESENTATION_ID = '1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI'

def extractIdOfSlideNum(pageElements):
    for element in pageElements:
	if element['shape']['text']['textElements'][1]['autoText']['type'] == u'SLIDE_NUMBER':
	    return element['objectId']


def main():
    """Shows basic usage of the Slides API.
    Prints the number of slides and elments in a sample presentation.
    """
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('slides', 'v1', http=creds.authorize(Http()))

    # Call the Slides API
    presentation = service.presentations().get(
        presentationId=PRESENTATION_ID).execute()
    slides = presentation.get('slides')

    print('The presentation contains {} slides:'.format(len(slides)))
    for i, slide in enumerate(slides):
        print('- Slide #{} contains {} elements.'.format(
            i + 1, len(slide.get('pageElements'))))
	# See if we can find the element that includes the page number
	pageNumId = None
	try:
	    pageNumId = extractIdOfSlideNum(slide.get('pageElements'))
	except:
	    pass
	print(pageNumId)

if __name__ == '__main__':
    main()
