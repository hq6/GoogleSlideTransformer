# Google Slides Transformation Tool

This tool (still under development) is intended to make it easier to automate
certain tasks in Google slides, but allowing one to programmatically select
certain elements and perform simple transformations on them in batch.

For example, if one wanted to update the font size on all page numbers, and not
all pages used the same master, one could do so automatically with the
following command.

    python SlideTransformer.py  '1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI' --filter 'type=SLIDE_NUMBER'  -t 'textStyle:fontSize=20'


The author is currently implemented transformations and filters in a rather ad
hoc manner, depending on what is needed for his particular presentation, and
would appreciate contributions.

## Current List of Supported Filters

 * `Key = Value`: Match PageElements whose nested dictionary representation
   contains at some level of the nesting the given key asssociated with the
   given value.

## Current List of Supported Transformations

 * `textStyle`
    * `fontSize`


## Setup Instructions

Given that this app is deployed as a stand-alone script rather than hosted on a web server, the author cannot provide client keys and secrets. Therefore, users must visit the [Google Slides API page](https://developers.google.com/slides/quickstart/python) and perform the following steps for initial setup.

1. Click on "Enable The Google Slides API"
2. Create a project.
3. Click "Next"
4. Click "Download Client Configuration" and save credentials.json in the
   directory this application will be executed from.
