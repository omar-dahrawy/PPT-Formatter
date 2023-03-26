import java.awt.Color;
import java.io.*;
import java.util.List;
import java.util.Random;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

// importing Apache POI environment packages
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.VerticalAlignment;

public class FormatTextPPT {
	
	private File originalFile;
	
	public  FormatTextPPT() throws IOException {
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("I Will Marry You .. I Swear");
		int returnValue = fileChooser.showDialog(null, "Choose");
		
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			originalFile = fileChooser.getSelectedFile();
			preparePPTFiles();
		} else if (returnValue == JFileChooser.CANCEL_OPTION) {
			System.exit(0);
		}
	}
	
	public void preparePPTFiles() throws IOException {
		
		//load ppt file
		FileInputStream inputStream = new FileInputStream(originalFile);
		XMLSlideShow ppt = new XMLSlideShow(inputStream);
		
		//get list of slides
		List<XSLFSlide> slides = ppt.getSlides();
		
		//randomize slides order
		shuffleSlides(ppt, slides);
		shuffleSlides(ppt, slides);
		shuffleSlides(ppt, slides);
		
		// numbering slides after shuffle starting from ascending 
		renumberSlides(ppt, slides);
		
		// duplicate all slides and place each duplicate slide before its original
		duplicateSlides(ppt, slides);
		
		// for each pair of the same slide, remove the text format of the correct answer in the first slide only
		modifyAnswers(ppt, slides);
				
		// set the text alignment for the question textbox to middle in each duplicate slide
		fixAlignmets(ppt, slides);
		
		// set the answers bullet style in each duplicate slide to match that of the original slide 
		fixBullets(ppt, slides);

		// export the newly created PowerPoint files
		exportNewFiles(ppt);
	}
	
	public void shuffleSlides(XMLSlideShow ppt ,List<XSLFSlide> slides) {
		
		for(int i=0 ; i<slides.size() ; i++) {
			XSLFSlide slide = slides.get(i);
			int rndm = new Random().nextInt(slides.size());
			ppt.setSlideOrder(slide, rndm);
		}
		System.out.println("Shuffled slides order");
	}
	
	public void renumberSlides(XMLSlideShow ppt ,List<XSLFSlide> slides) {
		
		for (XSLFSlide slide : slides) {
			XSLFTextShape title = slide.getPlaceholder(0);
			XSLFTextParagraph paragraph = title.getTextParagraphs().get(0);
			XSLFTextRun text = paragraph.getTextRuns().get(0);
			String titleText = text.getRawText();
			int indexOfDash = titleText.indexOf('-');
			titleText = titleText.substring(indexOfDash, titleText.length());
			titleText = slides.indexOf(slide)+1 + titleText;
			text.setText(titleText);
		}
		System.out.println("Renumbered slides");
	}

	public void duplicateSlides(XMLSlideShow ppt ,List<XSLFSlide> slides) {
		
		int size = slides.size();
		for (int i=0 ; i<size ; i++) {
			XSLFSlide slide = slides.get(i);
			XSLFSlide newSlide = ppt.createSlide();
			newSlide.clear();
			newSlide.appendContent(slide);
		}
		
		int j = 0;
		for(int i=slides.size()/2 ; i<slides.size() ; i++) {
			XSLFSlide slide = slides.get(i);
			ppt.setSlideOrder(slide, j);
			j+=2;
		}
		System.out.println("Duplicated slides");
	}

	public void modifyAnswers(XMLSlideShow ppt ,List<XSLFSlide> slides) {

		for (int i=0 ; i<slides.size() ; i+=2) {
			XSLFSlide slide = slides.get(i);
			XSLFTextShape title = slide.getPlaceholder(1);

			List<XSLFTextParagraph> paragraphs = title.getTextParagraphs();

			for (XSLFTextParagraph paragraph : paragraphs) {
				for (XSLFTextRun textRun : paragraph.getTextRuns()) {
					textRun.setUnderlined(false);
					textRun.setFontColor(Color.BLACK);
				}
			}
		}
		System.out.println("Modified answers");
	}
	
	public void fixAlignmets(XMLSlideShow ppt ,List<XSLFSlide> slides) {
		for (XSLFSlide slide : slides) {
			slide.getPlaceholders()[0].setVerticalAlignment(VerticalAlignment.MIDDLE);
		}
		System.out.println("Fixed text alignment");
	}
	
	public void fixBullets(XMLSlideShow ppt ,List<XSLFSlide> slides) {
		
		for (int i=0 ; i<slides.size() ; i+=2) {
			List<XSLFTextParagraph> paragraphs = slides.get(i).getPlaceholder(1).getTextParagraphs();
			for (int j=0 ; j<paragraphs.size() ; j++) {
				XSLFTextParagraph paragraph = paragraphs.get(j);
				paragraph.setBulletAutoNumber(slides.get(i+1).getPlaceholder(1).getTextParagraphs().get(j).getAutoNumberingScheme(), j+1);
			}
		}
		System.out.println("Fixed bullets numbering");
	}
	
	public void exportNewFiles(XMLSlideShow ppt) throws IOException {
		File pptWithAnswers = new File(originalFile.getParent() + "/NEW With Answers.pptx");
		File pptWithoutAnswers = new File(originalFile.getParent() + "/NEW Without Answers.pptx");
		FileOutputStream outWithAnswers = new FileOutputStream(pptWithAnswers);
		FileOutputStream outWithoutAnswers = new FileOutputStream(pptWithoutAnswers);
		
		ppt.write(outWithAnswers);
		outWithAnswers.close();
		
		for(int i=ppt.getSlides().size()-1 ; i>0 ; i-=2) {
			ppt.removeSlide(i);
		}
		
		ppt.write(outWithoutAnswers);
		outWithoutAnswers.close();
		
		ppt.close();
				
		String[] buttons = {"I Love Omar", "I Love Omar Gedan"};
	    JOptionPane.showOptionDialog(null, "Done :*", "For Salomti", JOptionPane.PLAIN_MESSAGE, 0, null, buttons, null);
		System.out.println("\nNew PowerPoint files exported successfully.");
	}
	
	public static void main (String args[]) throws IOException {
		
		new FormatTextPPT();
		
	}
}
