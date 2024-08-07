.. _intro-start:

Introduction
============

This Word Macro formats and exports the current document to a clean PDF with only the 
right hand bars showing tracked changes.

.. _user-manual-start:


User Manual
===========

Installation:
------------------------------
The *.BAS file can be run or imported directly to the VB Developer window
through the MSWord Developer Tab.  For availablity across all documents
the macro needs to be included in the Normal template used.

CAUTION:
------------------------------
The macro exports the PDF to the current directory of the file being edited.
The file needs to have been saved at least once to either a local drive or
to an O365 cloud folder. 
Macro can only be used on a document opened on the desktop app.  Cannot be used
within the O365 cloud editor.

Instructions:
------------------------------
The macro can be run directly from the VB window or assigned to a button within the ribbon.
The macro will check and determine if the file is local or cloud based and export it as needed.  
The macro does not edit the current document and only makes the changes to a temporary document 
that is opened at the time of export.  This preserves all comments and revision mark settings 
of the original document.

Features:
- Removes all formatting highlights and leaves only right side red revision bars
- Removes all comments
- Refreshes all document links, references, and tables without adding unneeded formatting marks
    or rev bars.
- Fixes double space after periods to single space
- Exports PDF to either the local folder or to the cloud folder the document is opened from.
