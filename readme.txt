First off, don't be rude, VOTE.
Even if you give it a one star, let me know what you think.

RTF Label
See project1.vbp for an example of usage and code.

I created this because I am making a WYSIWYG editor and there is NOTHING on PSC that does this. I will admit there are some atempts but nothing I have seen comes close to this.

Bascially this control BitBlts the form image under it into a picture box, then it BitBlts the rich text over it.

Known Issues....
I have to use the variable SID to tell the difference between controls. The hWND and HDC always return the same number.

When moving the control there is a horrible flashing, I am working on this, basically the next version will draw all graphics off screen (double buffering)

Images in the RFT file sometimes do not show up in display mode.

The RichTextBox will not send the image of the RTF while word wrapped. I have no idea why, but I am digging in the microsoft documentation for it.

If you plan on using this control in a project that uses a lot of them, you might want to modify it to remove the RTFbox from the control and link it to an external RTFbox, otherwose you end up with one RTFbox for each RTFlabel - and that will chew up your memory.

Special Thanks go to...
Andrew Heinlein - for the tranparent BitBlt
Unknown Author - for the VirtualDC class

Bryan Cairns
cairnsb@html-helper.com



