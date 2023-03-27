# PowerPoint Formatter
## Java program to automate formatting of PowerPoint files.

I created this project as a challenge for myself to automate a boring, systematic task that I had to do regularly.
The task consisted mainly of reformatting PowerPoint files: duplicating and reordering slides, modifying text style and formatting, fixing text alignment and bullet numbering. The steps are explained in detail in the **File modification steps** section in order to make understanding the code easier.

## Reuse instructions and External libraries
For reading and modifying PowerPoint files using Java, I used [Apache POI](https://poi.apache.org), a Java API for Microsoft Documents. In order to reuse the API for this project, simply download and unzip the [release artifacts](https://poi.apache.org/download.html) and add the .jar files to the project's classpath.

Once done, simply run the program and select the [sample PowerPoint file] using the file browser. The program will run its functions on the file and export two newly created .ppt files in the original file's directory, not affecting the original file.

## File modification steps
***This project was created for a very specific use case and it's highly unlikely that someone else will need to use it in its current implementation. However, understanding the steps executed will help you modify the program to match your needs.***

The [sample file]() contains 20 slides and each slide contains two text placeholders, with some slides containing some images. The first placeholder contains the question text, and all questions are numbered using the numbering format `1- ...`. The second placeholder contains the multiple choices, also numbered using the numbering format `a), b), c) ...`. In each slide, the correct answer is underlined and formatted with blue font color.

The goal of the task is to:
1. Shuffle the order of the slides
2. Renumber the questions starting from 1 ascending
3. Duplicate each slide, placing the duplicate slide right after the original
4. For each pair of slides, remove the text formatting for the correct answer from the first slide. This way, one slide has the question and the answers and the next slide also has the question and the answers - but with the correct answer underlined and colored.

At this point, the main steps of the task are completed. But due to how the Apache POI API works, the text alignment and the bullet style of the duplicated slides need fixing. Due to this, two more steps are needed:
5. Fix the vertical text alignment of the question text in each duplicate slide
6. Fix the bullet style of the answer choices text in each duplicate slide

These steps are executed in the same order by the program. After that, the program exports two new PowerPoint files. One file contains the final result of the above mentioned steps (one slide containing the question and the choice answers, and another duplicate slide with the correct answer underlined and colored). The other file contains only the slides with the question and choice answers.

With the above mentioned steps and the comments provided in the code, the program can be modified to format any PowerPoint files according to different use cases.
