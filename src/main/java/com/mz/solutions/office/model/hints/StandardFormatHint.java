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
package com.mz.solutions.office.model.hints;

import com.mz.solutions.office.extension.ExtendedValue;

import java.util.Objects;

/**
 * Hinweis an den Ersetzungsprozess einfache Layout/Formatierung auszuführen anstelle den
 * Platzhalter mit einem sichtbaren Wert zu belegen, zB um Absätze oder Tabellen zu entfernen.
 * 
 * <p>Nicht unterstützte Formatierungs-Hinweise werden als leere Zeichenkette ersetzt und führen
 * nicht zum Fehler oder Abbruch des Ersetzungsvorganges.</p>
 * 
 * @author Moritz Riebe
 */
public final class StandardFormatHint extends ExtendedValue {
    
    /**
     * Absatz wird weiterhin im Dokument belassen und nicht entfernt.
     * 
     * <p>Der Absatz wird vom Ersetzungsprozess nicht entfernt.</p>
     */
    public static final StandardFormatHint PARAGRAPH_KEEP = new StandardFormatHint("PARAGRAPH_KEEP");
    
    /**
     * Absatz wird beim Ersetzungsvorgang aus dem Dokument entfernt.
     * 
     * <p>Der Absatz in welchem dieser Platzhalter auftaucht wird vollständig im Ersetzungsprozess
     * aus dem Dokument entfernt. Nur der Absatz, in dem der Platzhalter steht. Dazu zählen auch
     * Inhalte von Tabellen-Zellen; wird dieser Formatierungs-Hinweis in einer Tabellen-Zelle
     * gefunden, wird der Inhalt der Zelle entfernt, aber die Zelle und Tabelle bleiben.</p>
     */
    public static final StandardFormatHint PARAGRAPH_REMOVE = new StandardFormatHint("PARAGRAPH_REMOVE");
    
    /**
     * Absatz wird als Versteckt/Hidden formatiert und beim Druck nicht angezeigt.
     * 
     * <p>Ändert die gesamte Formatierung des Absatzes auf "Hidden Text". Der Text, und damit der
     * Absatz, bleiben im bestehenden Dokument weiterhin vorhanden und wird im Ersetzungsprozess
     * auch ersetzt, sobald weitere Platzhalter dort enthalten sind.</p>
     * 
     * <p>Versteckter Text wird bei der Voreinstellung von MS Office und Open Office dem Benutzer
     * während des Bearbeitens angezeigt, aber nicht gedruckt oder exportiert (z.B. PDF). Abweichend
     * kann vom Nutzer eingestellt werden auch versteckten Text zu drucken/exportieren.</p>
     * 
     * <p>Diese Formatierung ist nicht zu verwechseln mit {@link #PARAGRAPH_REMOVE}, bei der der
     * vollständige Absatz völlig aus dem Dokument entfernt wird.</p>
     */
    public static final StandardFormatHint PARAGRAPH_HIDDEN = new StandardFormatHint("PARAGRAPH_HIDDEN");
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    /** Tabelle wird weiterhin im Dokument belassen und nicht entfernt. */
    public static final StandardFormatHint TABLE_KEEP = new StandardFormatHint("TABLE_KEEP");
    
    /** Tabelle wird beim Ersetzungsvorgang vollständig aus dem Dokument entfernt. */
    public static final StandardFormatHint TABLE_REMOVE = new StandardFormatHint("TABLE_REMOVE");
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private final String internIdentifier;

    private StandardFormatHint(String internIdentifier) {
        this.internIdentifier = internIdentifier.intern();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int hashCode() {
        return Objects.hashCode(this.internIdentifier);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean equals(Object otherObject) {
        if ((otherObject instanceof StandardFormatHint) == false) {
            return false;
        }
        
        if (getClass() != otherObject.getClass()) {
            return false;
        }
        
        final StandardFormatHint otherFormatHint = (StandardFormatHint) otherObject;
        return this.internIdentifier == otherFormatHint.internIdentifier
                || this.internIdentifier.equals(otherFormatHint.internIdentifier);
    }
    
    /**
     * {@inheritDoc}
     */
    @Override
    public String altString() {
        return "";
    }
    
}
