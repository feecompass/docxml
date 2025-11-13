header and footer references properties in section cause problems in turnaround test
some mixup with relationships between original docx (from Archive) and saved back to a new file (double relationship for a file in the output)

TextBox (wps:wsp) bodyPr currently supports only lIns, rIns, tIns, bIns, and anchor.
Unsupported bodyPr attributes include: wrap, vert, rtlCol, numCol, spcCol, colSz, fromWordArt, etc.
Only solid fills (a:solidFill/a:srgbClr) are supported for TextBox spPr; gradients/patterns not supported yet.
