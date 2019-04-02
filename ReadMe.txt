
The CRichEditCtrl turns out to be a bit unfriendly as far as concatenating RTF docs.
The streamer works on insertion, so it is incapable of adding more content to the end of a doc.
There may be some flag, but I could not find it in RichEdit.h.
What happens is the font characteristics  from the end of the doc that is being pasted to,
is set at the end of the doc after the insertion. like:

Control text, say red.
Inserted text, say black.

Result, if the user types into the bottom of the doc, the font size/color will follow like red.

It will be from the text above the insertion. Even if you SetSel(-1,-1).
The best I could think of was a parser hack on the RTF text.
It can be seen in the call: AppendRichText(...) AppendRichText
It assumes you want to end the document with paragraph control word.

Another thing I learned, and I certainly could not find it in the 'Word2007RTFSpec9',
that if you want to end a doc with a paragraph marker,
you must use two \par control words at the end of the doc.
CRichEditCtrl, Word, Open Office all add two \par in the case of ending with a \par.
