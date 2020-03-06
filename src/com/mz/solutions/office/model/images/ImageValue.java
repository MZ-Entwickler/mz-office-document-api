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

import com.mz.solutions.office.extension.ExtendedValue;
import java.util.Objects;
import java.util.Optional;
import javax.annotation.Nullable;

/**
 * Platzhalter-Wert {@link ExtendedValue} eines Bildes mit einer zugeordneten Resource
 * (Bild-Resource {@link ImageResource}) und ggf erweiterten Parametern/Einstellungen zum Bild wie
 * Abmaßhungen, Titel und Beschreibung (Alt-Text).
 * 
 * <p>Ein {@link ImageValue} entspricht einer Kombination aus einer Bild-Resource
 * ({@link ImageResource}) mit den erweiterten Angaben der Abmaßungen und optional Titel und
 * Beschreibung des Bildes. Werden keine Abmaßungen angeben, wird von {@code 3cm x 1cm} ausgegangen,
 * was nicht bedeutet das jedes zu ersetztende Bild diese Abmaßung automatisch bekommt.</p>
 * 
 * <pre>
 *  final ImageResource imageData = ImageResource.loadImage(
 *          Paths.get("image.png"), StandardImageResourceType.PNG);
 * 
 *  final ImageValue imageValue = new ImageValue(imageData)
 *          .setDimension(1.0, 1.0, UnitOfLength.CENTIMETERS)   // default 3cm x 1cm
 *          .setOverwriteDimension(true)                        // default false
 *          .setTitle("Image Title")                            // optional, default nothing
 *          .setDescription("Image Description/Alt-Text");      // optional, default nothing
 * 
 *  // DataValue erzeugen um benannte Text-Platzhalter oder Bild-Platzhalter zu erzuegen
 *  final DataValue value1 = new DataValue("IMAGE_1", imageValue);
 *  final DataValue value2 = new DataValue("IMAGE_2", imageValue);
 * 
 *  // Ein Text-Platzhalter namens 'IMAGE_1' sowie 'IMAGE_2' wird nun anstelle von Text mit dem
 *  // geladenen Bild ersetzt.
 *  // Existiert im Vorlagen-Dokument ein Bild, in dessen Eigenschaften [Name, Titel, Beschreibung,
 *  // Alt-Text] einer der beiden Platzhalter-Namen auftaucht, wird jenes Bild durch das selbst
 *  // geladene Bild ersetzt.
 * </pre>
 * 
 * <p><b>Einsetzen/Einfügen eines Bildes an der Stelle eines Text-Platzhalters:</b> Liegt ein
 * Text-Platzhalter ({@code MergeField || User-Def-Field}) vor, und an jener Stelle soll ein Bild
 * eingesetzt werden. Wird dies mit einem Als-Zeichen-Anker ({@code as-char}) versehen und an jener
 * Stelle im Text/Dokument eingesetz mit den in {@link #setDimension(double, double, UnitOfLength)}
 * übergeben Abmaßungen.</p>
 * 
 * <p><b>Ersetzen eines bestehenden Bildes als Bild-Platzhalter:</b> Bestehende Bilder in Dokumenten
 * können durch andere Bilder ersetzt werden. Zur Erkennung des Bildes, und dessen zugehörigem
 * Platzhalter-Namen, werden die Felder {@code TITLE/DESCRIPTION} bzw. {@code Alt-Text} verwendet.
 * Die Abmaßung des Bildes werden nicht geändert, es erfolgt nur ein Austausch des Bildes. Soll das
 * auszutauschende Bild auch neue Abmaßungen bekommen, müssen diese mit
 * {@link #setDimension(double, double, UnitOfLength)} angegeben werden und die Option
 * {@link #setOverwriteDimension(boolean)} muss auf {@link Boolean#TRUE} gesetzt werden. Ist die
 * Option {@code FALSE} werden bestehende Abmaßungen beibehalten und nicht überschrieben. Ist die
 * Option {@code TRUE} werden bestehende Abmaßungen auf die übergebenen Werte angepasst.</p>
 * 
 * <p><b>Titel, Beschreibung, Alt-Text:</b> Anhand jener Felder wird erkannt ob es zum Bild im
 * Dokument einen zugehörigen Platzhalter-Wert ({@link ImageValue}) gibt. Bei der Ersetzung können
 * jene Felder mit eigenen Werten überschrieben werden. Die Methoden {@link #setTitle(CharSequence)}
 * und {@link #setDescription(CharSequence)} sind jedoch optional.</p>
 * 
 * <p><b>Fehlende Unterstützung von Bildern:</b> Wird von der Office Implementierung ein Fall nicht
 * unterstützt, dann verhält sich {@link ImageValue} standard-konform wie {@link ExtendedValue}
 * und gibt die Bild-Beschreibung als alternativen Text in {@link #altString()} zurück.</p>
 * 
 * <p>Ein {@link ImageValue} kann mehrfach verwendet werden. Eine Bild-Resource
 * ({@link ImageResource}) kann gleichzeitig in mehreren {@link ImageValue}-Instanzen geteilt
 * verwendet werden um Platz und Resourcen zu im Ziel-Dokument zu sparen und kleinere Dokumente
 * erzeugen zu köennen.</p>
 * 
 * @see ImageResource
 *      Erzeugeten/Laden von Bild-Resourcen.
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public final class ImageValue extends ExtendedValue {

    private final ImageResource imgResource;
    
    @Nullable
    private String title, description;
    
    /** Flag ob bestehende Abmaßungen überschrieben werden sollen. */
    private boolean overwriteDimension = false;
    
    /** Breite des Bildes in Millimeter. */
    private double width = UnitOfLength.CENTIMETERS.toMillimeters(3.0D);
    
    /** Höhe des Bildes in Millimeter. */
    private double height = UnitOfLength.CENTIMETERS.toMillimeters(1.0D);
    
    /**
     * Erzeugt einen Platzhalter-Wert mit übergebener Bild-Resource.
     * 
     * @param imgResource   Bild-Resource, nie {@code null}, kann ebenfalls auch eine
     *                      lokale oder externe Resource sein.
     */
    public ImageValue(ImageResource imgResource) {
        this.imgResource = Objects.requireNonNull(imgResource, "imgResource");
    }
    
    /**
     * Rückgabe der zugewiesenen Bild-Resource.
     * 
     * @return  Resource, nie {@code null}.
     */
    public ImageResource getImageResource() {
        return imgResource;
    }
    
    /**
     * Setzen des optionalen Bild-Titels.
     * 
     * <p>Die Voreinstellung enthält keinen Bild-Titel.</p>
     * 
     * @param title         Bild-Titel, wenn {@code null} dann kein Titel.
     * 
     * @return              Eigene Instanz
     */
    public ImageValue setTitle(@Nullable CharSequence title) {
        this.title = (null == title) ? null : title.toString();
        return this;
    }
    
    /**
     * Setzen der optionalen Bild-Beschreibung.
     * 
     * @param description   Bild-Beschreibung, wenn {@code null} dann keine Beschreibung.
     * 
     * @return              Eigene Instanz
     */
    public ImageValue setDescription(@Nullable CharSequence description) {
        this.description = (null == description) ? null : description.toString();
        return this;
    }
    
    /**
     * Setzen der Bild-Abmaßungen für das Ziel-Dokument.
     * 
     * <p>Die übergebenen Abmaßungen finden nur Anwendung, wenn das Bild an die Stelle eines
     * normalen Text-Platzhalters eingesetzt/eingefügt wird oder wenn die Option
     * {@link #setOverwriteDimension(boolean)} auf {@link Boolean#TRUE} gesetzt wurde.</p>
     * 
     * <p>Als Voreinstellung sind {@code 3cm x 1cm} vergeben.</p>
     * 
     * @param width     Bild-Breite, in der mit {@code unit} angegeben Maß-Einheit.
     * 
     * @param height    Bild-Höhe, in der mit {@code unit} angegeben Maß-Einheit.
     * 
     * @param unit      Längen-Einheit für {@code width} und {@code height}.
     * 
     * @return          Eigene Instanz
     * 
     * @throws  IllegalArgumentException
     *          Wenn die übergebenen Abmaßungen kleiner oder gleich {@code 0} sind.
     * 
     * @throws  NullPointerException
     *          Wenn der Parameter {@code unit == null} ist.
     */
    public ImageValue setDimension(double width, double height, UnitOfLength unit) {
        Objects.requireNonNull(unit, "unit");
        
        if (width <= 0.0D) {
            throw new IllegalArgumentException("width <= 0.0");
        }
        
        if (height <= 0.0D) {
            throw new IllegalArgumentException("height <= 0.0");
        }
        
        this.width = unit.toMillimeters(width);
        this.height = unit.toMillimeters(height);
        
        return this;
    }
    
    /**
     * Setzen der Option bestehende Abmaßungen zu ignorieren und mit den eigenen Werten
     * zu überschreiben.
     * 
     * <p>Findet Anwendung wenn ein bestehendes Bild im Dokument (Bild-Platzhalter) durch ein
     * anderes Bild ersetzt werden soll. Wird {@code true} übergeben, wird nicht nur das Bild durch
     * das eigenen ersetzt, sondern die Abmaßungen werden ebenfalls durch die eigenen Angaben
     * abgewandelt.</p>
     * 
     * @param overwriteExistingDimension    Überschreiben/Ersetzen bestehender Abmaßungen.
     * 
     * @return          Eigene Instanz
     */
    public ImageValue setOverwriteDimension(boolean overwriteExistingDimension) {
        this.overwriteDimension = overwriteExistingDimension;
        return this;
    }
    
    /**
     * Rückgabe des optionalen Bild-Titels.
     * 
     * @return  Bild-Titel
     */
    public Optional<String> getTitle() {
        return Optional.ofNullable(title);
    }

    /**
     * Rückgabe der optionalen Bild-Beschreibung.
     * 
     * @return  Bild-Beschreibung
     */
    public Optional<String> getDescription() {
        return Optional.ofNullable(description);
    }
    
    /**
     * Rückgabe der gewünschten Bild-Breite in Millimeter.
     * 
     * @return  Breite in Millimeter.
     */
    public double getWidth() {
        return this.width;
    }
    
    /**
     * Rückgabe der gewünschten Bild-Höhe in Millimeter.
     * 
     * @return  Höhe in Millimeter
     */
    public double getHeight() {
        return this.height;
    }

    /**
     * Rückgabe der Option bestehende Bild-Abmaßungen zu überschreiben/ersetzen.
     * 
     * @return  wenn {@code true}, sollen bestehende Abmaßungen im Vorlagen-Dokument ignoriert
     *          und mit den eigenen Werten überschrieben werden.
     */
    public boolean isOverwriteDimension() {
        return overwriteDimension;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String altString() {
        return getDescription().orElse(null);
    }
    
}
