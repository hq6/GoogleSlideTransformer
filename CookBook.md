# Recipes for Transformations


Since the API is so verbose, the commands must also be verbose so as to be specific enough.

    # Change  outLine color of each ellipse with a themeColor for any of its color properties.
    python SlideTransformer.py -s '6'  -f 'shapeType:v=="ELLIPSE"' -f'color:"themeColor" in v' -t 'shapeProperties:outline.outlineFill.solidFill.color:themeColor=DARK2' 1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI

    # Update font size to 20 for all page numbers
    python SlideTransformer.py  -f 'type:v=="SLIDE_NUMBER"' -t 'textStyle:fontSize=20'  1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI

    # Match the size and position of a reference object; useful for aligning
    # figure builds across slides.
    python SlideTransformer.py  -s 32-34 -f 'image:True' -t 'transformPageElement:=&g4361db93fb_1_34'  '1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI'
