package mergedocx;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.docx4j.Docx4J;
import org.docx4j.dml.CTBlip;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageBmpPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageEpsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageGifPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageJpegPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImagePngPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageTiffPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.relationships.Relationship;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class MergeDocx {

	// https://stackoverflow.com/questions/23796468/merge-worddocx-documents-with-docx4j-how-to-copy-images

	private static final Logger logger = LoggerFactory.getLogger(MergeDocx.class);
	final static String[] sourceDocxNames = { "cover-header.docx", "simple1.docx", "simple3.docx", "simple2.docx" };
	static String outputfilepath = System.getProperty("user.dir") + "/target/out/merged.docx";

	public static void main(String[] args) {
		try {
			// Create list of docx packages to merge
			List<WordprocessingMLPackage> wmlPkgList = new ArrayList<>();

			// Load all source DOCX files into WordprocessingMLPackage objects
			for (String filename : sourceDocxNames) {
				logger.info("Loading " + filename);
				wmlPkgList.add(WordprocessingMLPackage.load(MergeDocx.class.getResourceAsStream("/" + filename)));
			}

			Map<String, String> map = new HashMap<>();
			map.put("TITLE", "This is my title");

			// Set the first document as the target (where other documents will be merged)
			WordprocessingMLPackage pkgTarget = wmlPkgList.get(0);
			MainDocumentPart docTarget = pkgTarget.getMainDocumentPart();
			VariablePrepare.prepare(pkgTarget);
			for (SectionWrapper section : pkgTarget.getDocumentModel().getSections()) {
				if (section.getHeaderFooterPolicy() != null) {
					section.getHeaderFooterPolicy().getDefaultHeader().variableReplace(map);
					section.getHeaderFooterPolicy().getDefaultFooter().variableReplace(map);
				}
			}

			// Loop through remaining documents to merge into the target
			for (int i = 1; i < wmlPkgList.size(); i++) {
				logger.info("\n==============================\nMerging document: " + sourceDocxNames[i]);

				WordprocessingMLPackage pkgSource = wmlPkgList.get(i);

				VariablePrepare.prepare(pkgSource);
				pkgSource.getMainDocumentPart().variableReplace(map);

				List<?> body = pkgSource.getMainDocumentPart().getJAXBNodesViaXPath("//w:body", false);
				for (Object b : body) {
					List<?> filhos = ((org.docx4j.wml.Body) b).getContent();
					for (Object k : filhos)
						pkgTarget.getMainDocumentPart().addObject(k);
				}

				List<Object> blips = pkgSource.getMainDocumentPart().getJAXBNodesViaXPath("//a:blip", false);
				for (Object el : blips) {
					try {

						CTBlip blip = (CTBlip) el;
						RelationshipsPart parts = pkgSource.getMainDocumentPart().getRelationshipsPart();
						Relationship rel = parts.getRelationshipByID(blip.getEmbed());
						Part part = parts.getPart(rel);

						if (part instanceof ImagePngPart)
							System.out.println(((ImagePngPart) part).getBytes());
						if (part instanceof ImageJpegPart)
							System.out.println(((ImageJpegPart) part).getBytes());
						if (part instanceof ImageBmpPart)
							System.out.println(((ImageBmpPart) part).getBytes());
						if (part instanceof ImageGifPart)
							System.out.println(((ImageGifPart) part).getBytes());
						if (part instanceof ImageEpsPart)
							System.out.println(((ImageEpsPart) part).getBytes());
						if (part instanceof ImageTiffPart)
							System.out.println(((ImageTiffPart) part).getBytes());

						Relationship newrel = pkgTarget.getMainDocumentPart().addTargetPart(part,
								AddPartBehaviour.RENAME_IF_NAME_EXISTS);

						blip.setEmbed(newrel.getId());
						pkgTarget.getMainDocumentPart().addTargetPart(
								pkgSource.getParts().getParts().get(new PartName("/word/" + rel.getTarget())));

					} catch (Exception ex) {
						ex.printStackTrace();
					}
				}
			}

			// Save the merged document
			File out = new File(outputfilepath);
			out.getParentFile().mkdirs(); // Ensure the output directory exists
			pkgTarget.save(out);
			logger.info("Merged document saved to " + outputfilepath);

			File pdf = new File(out.getParentFile(), "merged.pdf");

			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(out);

			FileOutputStream outputStream = new FileOutputStream(pdf);
			Docx4J.toPDF(wordMLPackage, outputStream);

		} catch (Exception e) {
			logger.error("Error merging documents: " + e.getMessage(), e);
		}
	}

}