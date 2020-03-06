/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2020,   Moritz Riebe     (moritz.riebe@mz-entwickler.de)
 *                       Andreas Zaschka  (andreas.zaschka@mz-entwickler.de)
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 * 
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package com.mz.solutions.office.model.images;

import com.mz.solutions.office.MicrosoftImageResourceType;
import com.mz.solutions.office.OfficeDocumentFactory;
import com.mz.solutions.office.OfficeProperty;
import com.mz.solutions.office.OpenDocumentImageResourceType;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.net.URL;
import java.nio.file.Path;
import java.util.Objects;

/**
 * Bild-Resource die das eigentliche Bild kapselt sowie dessen Bild-Typ (Dateiformat).
 * 
 * <p>Eine einmal geladene Bild-Resource kann Dokument übergreifend mehrfach verwendet werden und
 * muss nicht für jeden {@link ImageValue} neu geladen werden.</p>
 * 
 * <p>Wird eine Bild-Resource einmal geladen und mehrfach innerhalb eines Dokumentes verwendet,
 * ohne jedes mal eine neue Bild-Resource zu erzeugen, wird jene Resource im Ergebnis-Dokument,
 * unabhängig der Abmaßungen und Einstellungen, nur einmal hinterlegt um kleinere Dokumente
 * erzeugen zu können.</p>
 * 
 * <p><b>Arten von Bild-Resourcen:</b> Die einfachen {@code #loadImage*} Methoden laden
 * Bild-Resourcen direkt (in den Arbeitsspeicher) und betten diese im Ziel-Dokument ein. Dazu gibt
 * es noch Lokale- und Externe-Resourcen. <i>Lokale Resourcen</i> (Datei im Dateisystem) werden im
 * Dokument nur verlinkt und angezeigt, aber nicht im Ziel-Dokument direkt eingebettet. Es wird auf
 * das Quell-Bild als Dateipfad verwiesen. Dazu ist die Methode
 * {@link #useLocalFile(Path, ImageResourceType)} zu verwenden. <i>Externe Resourcen</i> sind URLs
 * die, wie lokale Resourcen, nicht im Ziel-Dokument eingebettet werden, sondern mit der URL im
 * Dokument eingebunden werden und zum Zeitpunkt des Ladens/Öffnens im Office-Programm erst von
 * der gegebenen URL geladen werden. Dazu ist die Methode
 * {@link #useExternalFile(URL, ImageResourceType)} zu verwenden.
 * <b>Lokale- und Externe-Resourcen werden nur dann direkt als Link/Verlinkung im Dokument angegeben
 * wenn in der {@link OfficeDocumentFactory} die Einstellung
 * {@link OfficeProperty#IMG_LOAD_AND_EMBED_EXTERNAL} auf {@link Boolean#FALSE} gesetz wurde! Die
 * Voreinstellung ({@link Boolean#TRUE}) lädt die lokale/externe Resource von der Datei/URL und
 * bindet diese direkt im Ziel-Dokument ein.</b></p>
 * 
 * <p><b>Dummy-Farb-Bilder:</b> Zum Ersetzen und ggf. zum Testen/Debuggen können Bilder mit
 * 1x1 Pixel erzeugt werden und im Dokument verwendet werden. Mit {@link #dummyColorImage(int)} und
 * {@link #dummyColorImage(int, int, int)} können diese Bild-Resourcen erzeugt werden. Dabei wird
 * ein {@code BMP} Bild erzeuget das nur {@code 58 Byte} groß ist und der übergebenen
 * RGB Farbe entspricht.</p>
 * 
 * <p><b>Datei-Format der Bild-Resourcen ({@link ImageResourceType}):</b> Das Dateiformat der
 * geladenen/übergebenen Bild-Resource ist stets möglichst korrekt anzugeben. Zur Vereinfach sind in
 * {@link StandardImageResourceType} die wichtigsten Dateiformate angegeben. Um spezielle Formate
 * spezifischer Office-Implementierungen zu verwenden, können {@link OpenDocumentImageResourceType}
 * und {@link MicrosoftImageResourceType} verwendet werden. Die Schnittstelle
 * {@link ImageResourceType} kann auch selbst implementiert werden.</p>
 * 
 * <pre>
 *  final ImageResource image = ImageResource.loadImage(
 *          Paths.get("smiley.png"), StandardImageResourceType.PNG);
 * 
 *  final ImageValue imageLarge = new ImageValue(image)
 *          .setDimension(15.0, 15.0, UnitOfLength.CENTIMETERS)
 *          .setTitle("Image Large");
 * 
 *  final ImageValue imageSmall = new ImageValue(image)
 *          .setDimension(0.5, 0.5, UnitOfLength.CENTIMETERS)
 *          .setTitle("Image Small");
 * 
 *  final DataValue value1 = new DataValue("IMAGE_SMALL", imageSmall);
 *  final DataValue value2 = new DataValue("IMAGE_LARGE", imageLarge);
 * </pre>
 * 
 * @see StandardImageResourceType
 *      Wichtigte Bild-Formate die von beiden Office-Implementierungen unterstützt werden
 * 
 * @see OpenDocumentImageResourceType
 *      [Spezifisch] LibreOffice/OpenOffice Bild-Formate
 * 
 * @see MicrosoftImageResourceType
 *      [Spezifisch] Microsoft Office Bild-Formate
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public interface ImageResource {
    
    /**
     * Lädt eine Bild-Resource von der übergebenen Datei mit dem übergebenen Datei-Format sofort.
     * 
     * @param imageFile     Bild-Datei, nicht {@code null}.
     * 
     * @param formatType    Datei-Format, nicht {@code null}.
     * 
     * @return              Geladene Bild-Resource
     * 
     * @throws  UncheckedIOException 
     *          Gewrappte {@link IOException} wenn beim Laden der Datei ein Fehler auftrat.
     * 
     * @throws  NullPointerException
     *          Wenn einer der Parameter {@code null} ist.
     */
    public static ImageResource loadImage(Path imageFile, ImageResourceType formatType)
            throws UncheckedIOException
    {
        Objects.requireNonNull(imageFile, "imageFile");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.EagerEmbedImgResFile(imageFile, formatType);
    }
    
    /**
     * Lädt eine Bild-Resource direkt vom obergeben {@link InputStream} mit dem übergeben
     * Datei-Format ohne den {@link InputStream} zu schließen.
     * 
     * @param imageStream   Daten-Strom, wird nach vollständigem Lesen nicht geschlossen.
     * 
     * @param formatType    Datei-Format, nicht {@code null}.
     * 
     * @return              Geladene Bild-Resource
     * 
     * @throws  UncheckedIOException 
     *          Gewrappte {@link IOException} wenn beim Laden der Datei ein Fehler auftrat.
     * 
     * @throws  NullPointerException
     *          Wenn einer der Parameter {@code null} ist.
     */
    public static ImageResource loadImage(InputStream imageStream, ImageResourceType formatType)
            throws UncheckedIOException
    {
        Objects.requireNonNull(imageStream, "imageStream");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.EagerEmbedImgResData(imageStream, formatType);
    }
    
    /**
     * Erzeugt eine Bild-Resource aus dem übergebenen Byte-Array mit dem entsprechenden Bild-Format.
     * 
     * @param imageData     Bild-Daten
     * 
     * @param formatType    Bild-Format
     * 
     * @return              Daraus erzeugte Bild-Resource.
     * 
     * @throws  NullPointerException
     *          Wenn einer der Parameter {@code null} ist.
     */
    public static ImageResource loadImage(byte[] imageData, ImageResourceType formatType) {
        Objects.requireNonNull(imageData, "imageData");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.EagerEmbedImgResData(imageData, formatType);
    }
    
    /**
     * Erzeugt eine Bild-Resource die bei Bedarf von der übergebenen Datei geladen werden soll.
     * 
     * <p>Der Ladevorgang der Bild-Resource erfolgt erst beim erstmaligen Aufruf von
     * {@link ImageResource#loadImageData()} und wird darauf hin zwischengespeichert. Ist zu jenem
     * Zeitpunkt die Datei nicht ladbar, wird beim Ersetzungsvorgang eine
     * {@link UncheckedIOException} geworfen.</p>
     * 
     * @param imageFile     Bild-Datei
     * 
     * @param formatType    Bild-Format
     * 
     * @return              Bild-Resource die erst bei Bedarf von der übergeben Datei geladen wird.
     */
    public static ImageResource loadImageLazy(Path imageFile, ImageResourceType formatType) {
        Objects.requireNonNull(imageFile, "imageFile");
        Objects.requireNonNull(formatType, "formatType");
        
        return new StandardImageResource.LazyEmbedImgResFile(imageFile, formatType);
    }
    
    /**
     * Erzeugt eine Bild-Resouce anhand einer lokal verfügbaren Datei die, wenn unterstützt, von der
     * jeweiligen Office-Implementierung verlinkt wird und nicht im Ziel-Dokument eingebettet wird.
     * 
     * <p>Ist die Voreinstellung der {@link OfficeDocumentFactory} mit der Eigenschaft
     * {@link OfficeProperty#IMG_LOAD_AND_EMBED_EXTERNAL} nicht auf {@link Boolean#FALSE} geändert
     * worden, verhält sich diese Methode wie {@link #loadImageLazy(Path, ImageResourceType)}
     * und bettet die Bild-Resource im Ziel-Dokument ein.</p>
     * 
     * @param imageFile     Bild-Datei, sollte beim Öffnen des Ergebnis-Dokumentes vom jeweiligen
     *                      Office zugreifbar sein.
     * 
     * @param imgType       Bild-Format
     * 
     * @return              Lokale Bild-Resource
     */
    public static LocalImageResource useLocalFile(Path imageFile, ImageResourceType imgType) {
        Objects.requireNonNull(imageFile, "imageFile");
        Objects.requireNonNull(imgType, "imgType");
        
        return new StandardImageResource.LocalImageResourceImpl(imageFile, imgType);
    }
    
    /**
     * Erzeugt eine Bild-Resource anhand der übergeben URL die, wenn unterstützt, von der office-
     * Implementierung direkt als URL verlinkt wird und nicht im Ziel-Dokument eingebettet wird.
     * 
     * <p>Ist die Voreinstellung der {@link OfficeDocumentFactory} mit der Eigenschaft
     * {@link OfficeProperty#IMG_LOAD_AND_EMBED_EXTERNAL} nicht auf {@link Boolean#FALSE} geändert
     * worden, verhält sich diese Methode wie {@link #loadImageLazy(Path, ImageResourceType)}
     * und bettet die Bild-Resource im Ziel-Dokument ein.</p>
     * 
     * @param imageURL  URL der Bild-Resource
     * 
     * @param imgType   Bild-Format
     * 
     * @return          Externe Bild-Resource
     */
    public static ExternalImageResource useExternalFile(URL imageURL, ImageResourceType imgType) {
        Objects.requireNonNull(imageURL, "imageURL");
        Objects.requireNonNull(imgType, "imgType");
        
        return new StandardImageResource.ExternalImgResImpl(imageURL, imgType);
    }
    
    /**
     * Erzeugt eine Bild-Resource mit 1x1 Pixel in der entsprechend übergebenen Farbe im BMP-Format.
     * 
     * <pre>
     *  final ImageResource imageWhite = ImageResource.dummyColorImage(0xFFFFFF);
     *  final ImageResource imageBlack = ImageResource.dummyColorImage(0x000000);
     * 
     *  // Es kann auch java.awt.Color verwendet werden
     *  final ImageResource imageBlue = ImageResource.dummyColorImage(Color.BLUE.getRGB());
     * </pre>
     * 
     * @param rgbColor  RGB-Farbwert z.B. {@code 0xFFFFFF}.
     * 
     * @return          Bild-Resource, 1x1 Pixel in der entsprechenden Farbe mit einer größe von
     *                  58 Bytes.
     */
    public static ImageResource dummyColorImage(int rgbColor) {
        return dummyColorImage(
                (rgbColor >> 16) & 0xFF,
                (rgbColor >>  8) & 0xFF,
                (rgbColor >>  0) & 0xFF);
    }
    
    /**
     * Erzeugt eine Bild-Resource mit 1x1 Pixel in der entsprechend übergeben Farbe im BMP-Format.
     * 
     * <pre>
     *  final ImageResource imageWhite = ImageResource.dummyColorImage(0xFF, 0xFF, 0xFF);
     *  final ImageResource imageBlack = ImageResource.dummycolorImage(0x00, 0x00, 0x00);
     * 
     *  // Es kann auch java.awt.Color verwendet werden
     *  final ImageResource imageBlue = ImageResource.dummyColorImage(
     *          Color.BLUE.getRed(),
     *          Color.BLUE.getGreen(),
     *          Color.BLUE.getBlue());
     * </pre>
     * 
     * @param red   Rot-Anteilt, {@code 0x00..0xFF},
     * 
     * @param green Grün-Anteil, {@code 0x00..0xFF}
     * 
     * @param blue  Blau-Anteil, {@code 0x00..0xFF}
     * 
     * @return      Bild-Resource, 1x1 Pixel in der entsprechenden Farbe mit einer Größe
     *              von 58 Bytes.
     */
    public static ImageResource dummyColorImage(int red, int green, int blue) {
        final byte[] bmpImageData = new byte[] {
            (byte) 0x42, (byte) 0x4D, (byte) 0x3A, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x36, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x28, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x01, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x01, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x01, (byte) 0x00, (byte) 0x18, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x04, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0xC4, (byte) 0x0E, (byte) 0x00, (byte) 0x00,
            (byte) 0xC4, (byte) 0x0E, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00, (byte) 0x00,
            (byte) (blue < 0x00 ? 0x00 : (blue > 0xFF ? 0xFF : blue)),
            (byte) (green < 0x00 ? 0x00 : (green > 0xFF ? 0xFF : green)),
            (byte) (red < 0x00 ? 0x00 : (red > 0xFF ? 0xFF : red)), (byte) 0x00
        };
        
        return loadImage(bmpImageData, StandardImageResourceType.BMP);
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /**
     * Gibt das Bild-Format dieser Bild-Resource zurück.
     * 
     * @return  Bild-Format, darf nie {@code null} zurück geben.
     */
    public ImageResourceType getImageFormatType();
    
    /**
     * Gibt das Bild als Byte-Array zurück.
     * 
     * <p>Mehrmaliges Aufrufen dieser Methode sollte am Besten immer das selbe Array zurück geben
     * und nach einem Ladevorgang die Bild-Resource zwischenspeichern.</p>
     * 
     * @return  Bil-Daten als Array, nie {@code null}.
     */
    public byte[] loadImageData();
    
}
