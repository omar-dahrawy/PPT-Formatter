# PPT-Formatter
## Java program to automate formatting of PowerPoint files.

I created this project as a challenge for myself to automate a boring, systematic task that I had to do regularly.
The task consisted mainly of reformatting PowerPoint files: duplicating and reorering slides, modifying text format, fixing text alignment and bullet numbering. The steps are explained in detail in the last section in order to make understanding the code easier.

### Reuse Instructions and External libraries
For reading and modifying PowerPoint files using Java, I used [Apache POI](https://poi.apache.org), a Java API for Microsoft Documents. In order to reuse the API for this project, simply download and unzip the [release artifacts](https://poi.apache.org/download.html) and add the .jar files to the project's classpath.

Once done, simply run the program and select the [sample PowerPoint file] using the file browser. The program will run its functions on the file and export two newly created .ppt files in the original file's directory, not affecting the original file.

### File modification steps
**This project was created for a very specific use case and it's highly unlikely that someone else will need to use it in its current implementation. However, understanding the steps executed will help you modify the program to match your needs.**
