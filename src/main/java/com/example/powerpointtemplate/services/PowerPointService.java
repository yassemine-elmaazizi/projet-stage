package com.example.powerpointtemplate.services;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class PowerPointService {

  public void processPresentation() {
    try {
      // Charger le modèle depuis le classpath
      File templateFile = new File("src/main/resources/Resume_Template.pptx");
      InputStream inputStream = new FileInputStream(templateFile);
      XMLSlideShow ppt = new XMLSlideShow(inputStream);

      // Lire les données JSON
      File jsonFile = new File("src/main/resources/infoCV.json");
      ObjectMapper objectMapper = new ObjectMapper();
      Map<String, Object> data = objectMapper.readValue(jsonFile, Map.class);

      // Obtenir la première diapositive
      XSLFSlide slide = ppt.getSlides().get(0);

      // Collecter les formes à remplacer
      List<XSLFShape> shapesToReplace = new ArrayList<>();

      // Trouver et manipuler les formes
      for (XSLFShape shape : slide.getShapes()) {
        if (shape instanceof XSLFTextBox) {
          XSLFTextShape textShape = (XSLFTextShape) shape;
          // Récupérer le texte de la forme et remplacer les espaces réservés
          for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
              String text = textRun.getRawText();
              for (String key : data.keySet()) {
                text = text.replace("{{" + key + "}}", data.get(key).toString());
              }
              textRun.setText(text);
              // Réduire la taille de la police pour éviter le chevauchement
              textRun.setFontSize(textRun.getFontSize() - 2);
            }
          }
        } else if (shape instanceof XSLFTable) {
          XSLFTable table = (XSLFTable) shape;
          for (XSLFTableRow tableRow : table.getRows()) {
            for (XSLFTableCell tableCell : tableRow.getCells()) {
              for (XSLFTextParagraph textParagraph : tableCell.getTextParagraphs()) {
                for (XSLFTextRun textRun : textParagraph.getTextRuns()) {
                  String text = textRun.getRawText();
                  for (String key : data.keySet()) {
                    text = text.replace("{{" + key + "}}", data.get(key).toString());
                  }
                  textRun.setText(text);
                  // Réduire la taille de la police pour éviter le chevauchement
                  textRun.setFontSize(textRun.getFontSize() - 2);
                }
              }
            }
          }
        } else if (shape instanceof XSLFPictureShape) {
          XSLFPictureShape pictureShape = (XSLFPictureShape) shape;
          String altText = getAltText(pictureShape);
          for (int i = 0; i < 10; i++) {
            if (altText.equals("{{TECH_" + i + "}}")) {
              shapesToReplace.add(pictureShape);
            }
          }
        }
      }

      // Remplacer les formes collectées par des icônes spécifiques
      String[] logoPaths = {
              "src/main/resources/logos/aws.png",
              "src/main/resources/logos/azure.png",
              "src/main/resources/logos/cue.png",
              "src/main/resources/logos/django.png",
              "src/main/resources/logos/docker.png",
              "src/main/resources/logos/elk.png",
              "src/main/resources/logos/gcp.png",
              "src/main/resources/logos/git.png",
              "src/main/resources/logos/gitlab.png",
              "src/main/resources/logos/go.png",
              "src/main/resources/logos/java.png",
              "src/main/resources/logos/javascript.png",
              "src/main/resources/logos/jen",
              "src/main/resources/images/Image1.png",
              "src/main/resources/images/Image2.png",
              "src/main/resources/images/Image3.png",
              "src/main/resources/images/Image4.png",
              "src/main/resources/images/Image5.png",
              "src/main/resources/images/Image6.png",
              "src/main/resources/images/Image7.png",
              "src/main/resources/images/Image8.png",
              "src/main/resources/images/Image9.png"
      };

      for (int i = 0; i < shapesToReplace.size() && i < logoPaths.length; i++) {
        XSLFShape shapeToReplace = shapesToReplace.get(i);
        String logoPath = logoPaths[i];

        // Ajouter l'image à la diapositive
        FileInputStream logoInputStream = new FileInputStream(logoPath);
        byte[] logoBytes = new byte[logoInputStream.available()];
        logoInputStream.read(logoBytes);
        logoInputStream.close();

        // Ajouter l'image
        XSLFPictureData pictureData = ppt.addPicture(logoBytes, PictureData.PictureType.PNG);
        XSLFPictureShape newPictureShape = slide.createPicture(pictureData);

        // Positionner l'image à l'endroit de la forme d'origine
        newPictureShape.setAnchor(shapeToReplace.getAnchor());

        // Supprimer la forme à remplacer
        slide.removeShape(shapeToReplace);
      }

      // Enregistrer la présentation modifiée
      FileOutputStream outputStream = new FileOutputStream("output.pptx");
      ppt.write(outputStream);
      outputStream.close();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  private String getAltText(XSLFPictureShape pictureShape) {
    XmlObject xmlObject = pictureShape.getXmlObject();
    XmlCursor cursor = xmlObject.newCursor();
    cursor.selectPath("./*");
    while (cursor.toNextSelection()) {
      XmlObject obj = cursor.getObject();
      if (obj instanceof CTNonVisualDrawingProps) {
        CTNonVisualDrawingProps nvProps = (CTNonVisualDrawingProps) obj;
        if (nvProps.isSetDescr()) {
          return nvProps.getDescr();
        }
      }
    }
    return "";
  }
}
