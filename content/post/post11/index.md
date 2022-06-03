---
# Documentation: https://wowchemy.com/docs/managing-content/

title: "Work with the bibliography."
subtitle: "Everything you need to know about working with bibliography"
summary: ""
authors: [Volchok Kristina]
tags: []
categories: []
date: 2022-06-03T15:12:43+03:00
lastmod: 2022-06-03T15:12:43+03:00
featured: false
draft: false

# Featured image
# To use, add an image named `featured.jpg/png` to your page's folder.
# Focal points: Smart, Center, TopLeft, Top, TopRight, Left, Right, BottomLeft, Bottom, BottomRight.
image:
  caption: ""
  focal_point: ""
  preview_only: false

# Projects (optional).
#   Associate this post with one or more of your projects.
#   Simply enter your project's folder or file name without extension.
#   E.g. `projects = ["internal-project"]` references `content/project/deep-learning/index.md`.
#   Otherwise, set `projects = []`.
projects: []
---


The Word object model includes several objects designed to automate the creation of bibliographies. The following table lists the main objects of the Word Bibliography function. Use these objects and additional properties and methods in the Word object model to add sources to source lists, reference sources in a document, and manage sources.
Understanding the XML source
Sources are added to the source lists programmatically using XML strings. Depending on the type of source that needs to be added, the XML structure changes. To define the XML structure for a source type, you can add the same source type manually and then view the returned XML. The following describes how to do this.
The Guid and LCID elements are optional, but you can provide values for them if you wish. The value of the Guid element must be a valid GUID that can be created programmatically outside of the Word object model. (See Visual Studio documentation or Windows MSDN documentation for information about programmatically creating IDs.) Word creates GUIDs when users add or edit the source. If you don't add a GUID to the XML, and the user then edits the source, Word creates a GUID. This allows Word to determine which source is the most recent, based on the GUID value, and hints whether the Word user wants to update the outdated source to maintain continuity between the main list and the current list.

LCID specifies the source language. (See msDN for valid language identification values.) Word uses LCID to find out how to display the cited source in the document bibliography. For example, one source can be written in French, one in English and one in Japanese. From the LCID, Word determines how to display names (for example, Last, First for English), which punctuation to use (for example, using a comma in one language and a comma in another) and which strings to use (for example, whether to use "et al" or another localized form).

After removing the optional elements, there may be a structure similar to the following XML structure. (You can determine which elements are needed, since there is no corresponding editable field in the Create Source dialog box. Omitting one or more necessary elements causes an error during operation.)


Now that you have the basic XML source structure for the book, you can add additional book sources to the main source list and the current source list. Additional items can be found by checking the Show all Bibliography fields field.
Adding sources to the source data list and the current source list
Adding sources to the main source list is similar to adding sources to the current source list, except for accessing a collection of Sources from various main objects. To add a source to the source data list, you can access the Source collection from the Bibliography property of the Application object. To add a source to the current source list, you can access the Source collection from the bibliography property of the Document object.

In the following example, the basic structure defined earlier is used to add another book source to the main source list.
The string can be changed to Application.Bibliography.Sources.Add strXml``ActiveDocument.Bibliography.Sources.Add strXml

Including a source programmatically in the source code list does not automatically add it to the current source list. However, to add a quote to a document, the source must be listed in the current list of sources. You can manually copy one or more sources from the master list to the current list using the Source Manager dialog boxes or programmatically copy one or more sources from the main list to the current list. The following example copies all the sources in the master source to the current source. After the sources are added to the current list, links for these sources can be inserted into the document.
Share the original list
Sometimes it may be necessary to share the original list with others in the organization. When adding sources to the master list, Word adds them to the file names "sources.xml ", located at C:\Users \<user>\AppData\Roaming\Microsoft\Bibliography\sources.xml . You can share this file with other users by providing them with a file that users can download manually from the Source Manager dialog box or programmatically through code.


Inserting a citation
You can insert a bibliography citation using the Add method for the Fields collection. In the following example, a quote for the source added earlier is inserted in the cursor. The text for the field is equal to the tag value or the Tag element value, which in this case is "Mor01". (See the XML code in the AddBibSource subroutine shown earlier for the XML string "<b:Tag>Mor01</b:Tag>".) The value of the Tag element also corresponds to the Tag property for the Source object.

Applying the Bibliography style
After inserting the bibliography into the document, you can set the bibliography style. Word formats several different styles of bibliographies. The bibliography style can be set using the BibliographyStyle property. This property can be one of the following String values :

APA

Chicago

GB7714

GOST — sorting names

GOST — sorting of names

ISO 690 — date of the first element

ISO 690 — numerical reference

MLA

SISTO2

Turabian

Insert bibliography
As in citations, fields are used in bibliographies. To insert a bibliography, you must insert a field with the constant wdFieldBibliography specified for the field type. The following code inserts the bibliography into the active document on the cursor. In this example, it is assumed that the cursor is located at the end of the document or on a new page.
