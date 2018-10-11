# Google Slides Transformation Tool

This tool (under development) is intended to make it easier to automate
certain tasks in Google slides, but allowing one to programmatically select
certain elements and perform simple transformations on them in batch.

For example, if one wanted to update the font size on all page numbers, and not
all pages used the same master, one could do so automatically with the
following command.

    python SlideTransformer.py   --filter 'type:v=="SLIDE_NUMBER"'  -t 'textStyle:fontSize=8' '1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI'

If one wanted to update the fill color of a set of ellipses on a particular slide, one could use the following command.

    python SlideTransformer.py -s '6'  -f 'shapeType:v=="ELLIPSE"' -f'color:"themeColor" in v' \
        -t 'shapeProperties:shapeBackgroundFill.solidFill.color=themeColor:ACCENT4' \
        1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI


The author is currently implementing transformations and filters in a rather ad
hoc manner, depending on what is needed for his particular presentation, and
would appreciate contributions.

## Current List of Supported Filters

 * `Key = <lambda_using_v>`: Match PageElements whose nested dictionary representation
   contains at some level of the nesting the given key asssociated with a value
   satisfying the lambda. The lambda passes in the value associated with a key
   using the name `v`.

## Current List of Supported Transformations

 * `textStyle`
    * `fontSize`
 * `shapeProperties`
    * Use `shapeProperties:<key>=&<objectId>` to transform an arbitrary key's
      value to match that of another object.
 * `transformPageElement`
    * Use `'transformPageElement:=&g4361db93fb_1_34'` to match the transformation of another page element.


## Setup Instructions

Given that this app is deployed as a stand-alone script rather than hosted on a web server, the author cannot provide client keys and secrets. Therefore, users must visit the [Google Slides API page](https://developers.google.com/slides/quickstart/python) and perform the following steps for initial setup.

1. Click on "Enable The Google Slides API"
2. Create a project.
3. Click "Next"
4. Click "Download Client Configuration" and save credentials.json in the
   directory this application will be executed from.

## Dependencies

 * [Google API Python Client](https://developers.google.com/slides/how-tos/libraries#python).

```
 pip install --upgrade google-api-python-client
```
 * OAuth 2 Client
```
sudo pip install --upgrade oauth2client
```

## References

 * https://developers.google.com/slides/reference/rest/v1/presentations.pages/other#Page.ThemeColorType
 * https://developers.google.com/slides/reference/rest/v1/presentations/request
 * https://developers.google.com/slides/reference/rest/v1/presentations.pages#Page.PageElement
