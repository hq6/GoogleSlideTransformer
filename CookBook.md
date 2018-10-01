# Recipes for Transformations


Since the API is so verbose, the commands must also be verbose so as to be specific enough.

    # Transform the outLine color of each ellipse with a themeColor for any of its color properties.
    python SlideTransformer.py -s '6'  -f 'shapeType:v=="ELLIPSE"' -f'color:"themeColor" in v' -t 'shapeProperties:outline.outlineFill.solidFill.color=themeColor:DARK2' 1CItHgsDlA1Sbz_g4zXh678C5_qCiAtOf7mKVQH9rYiI

