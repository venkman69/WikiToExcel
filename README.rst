WikiToExcel
-----------


Use is trivial as shown below::

    from wikitoexcel import wikitoexcel
    w2e = wikitoexcel(infile="./wikitbl.txt")
    # print the html representation of wiki markup
    print w2e.getHTML()
    # save excel file to out.xlsx
    w2e.saveExcel(fileName="out.xlsx")

Options are::

    wikitoexcel(wikiContent=<wikistr>, infile=<wikiTextFile>)

Features
--------

wikitoexcel can capture:

- Font styling: bold, underline, strikethrough
- Cell styling: foreground color, background color
- Supports merged cells
- Sheet name is captured as caption of the wiki table
- Multiline cell contents are styled with 'wrap' in excel
- span and div elements are converted to their inner text representation
- HyperLinks are addressed in a way that the hyperlink and display text are 
presented side by side. This facilitates round-trip between exceltowiki and wikitoexcel. 
Such as:
  http://yahoo.com Yahoo!

wikitoexcel currently cannot capture anything more complex than the above list. 
Features such as a font styling that within a paragraph is not honoured.

Release Notes: 0.1.0
--------------------
Initial Release