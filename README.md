# Cy Document Renderer
## Introduction
This is a very simple project that takes a bunch of text files 
and turns them into a microsoft word docx file. There are many 
different ways to do this 
[that are more customisable already](https://products.aspose.com/pdf/python-net/conversion/tex-to-docx/).
The goals of this project, however, are entirely different.

* Keep the text files simple. This is because I often write on small screens (e.g. my phone)
and I want to avoid the need for lots of codes and commands to remember
* Avoid the need for [learning Latex](https://www.overleaf.com/learn/latex/Learn_LaTeX_in_30_minutes#:~:text=Writing%20your%20first%20piece%20of%20LaTeX,-The%20first%20step&text=You%20can%20do%20this%20on,a%20new%20project%20in%20Overleaf.&text=Open%20this%20example%20in%20Overleaf.&text=You%20can%20see%20that%20L,of%20that%20formatting%20for%20you.)
* The text files should be just as readable as the word docs they produce


## How to use
Create several text files in the same directory. Use the following structure:
* files named "chapter00.txt" are considered header pages (e.g. title page, foreward, etc). They will
be processed alphabetically, so naming them like "chapter00-A", "chapter00-B" is a good idea
* files named "chapter01.txt" and beyond are book chapters.
* Use tags to customise different things in the document

Make sure you copy the .env.sample file and replace the values inside with your book details

### Tags
You can use tags to change the formatting
* `[MT]` - Denotes a Main Title, or book title. Uses the Header1 style in word
* `[By]` - Denotes a by line or an author name. Usually placed below the title
* `[T]` - Denotes a chapter title
* `[S]` - Denotes a chapter subtitle
* `###` - (on a separate line) - Denotes a section break

An example of this used in a chapter definition may be a text file like the following:
```
[T] My Chapter Title
[S] The Subtitle
The first paragraph
The second paragraph
```

This will produce a page in a word document with "My Chapter Title" as the chapter title in a 
header format, and "The Subtitle" just underneath as the subtitle text. There will be two 
paragraphs after that as each line is treated as its own separate thing.
