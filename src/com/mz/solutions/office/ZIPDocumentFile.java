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
package com.mz.solutions.office;

import com.mz.solutions.office.OfficeDocumentException.InvalidDocumentFormatForImplementation;
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UncheckedIOException;
import static java.nio.charset.StandardCharsets.UTF_8;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import static java.util.stream.Collectors.toList;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;
import static mz.solutions.office.resources.MessageResources.formatMessage;

/**
 * Öffnet (im RAM!), klont und speichert (packt) ZIP Dateien unter den
 * besonderen Voraussetzungen von Office ZIP Dateien.
 * 
 * <p>Das Gesamte Dokument wird vollständig in den RAM geladen.
 * Kompressionsmethoden, Reihenfolgen, Dateinamen etc... werden alle soweit
 * wie möglich beibehalten.</p>
 * 
 * @date        2015-03-10
 * @author      Riebe, Moritz       (moritz.riebe@mz-entwickler.de)
 * @author      Zaschka, Andreas    (andreas.zaschka@mz-solutions.de)
 */
final class ZIPDocumentFile {
    
    private static final String MSG_MISSING_ITEM = "ZIPDocumentFile_MissingItem";
    
    // List<> weil Reihenfolge der Einträge WICHTIG ist!
    private final List<ZipFileItem> zipItems;
    
    public ZIPDocumentFile(final Path documentFile) {
        this.zipItems = new LinkedList<>();
        readZipFile(documentFile);
    }
    
    private ZIPDocumentFile(final List<ZipFileItem> copyOfZipItems) {
        this.zipItems = copyOfZipItems;
    }
    
    public ZIPDocumentFile cloneDocument() {
        final List<ZipFileItem> clonedItems = new LinkedList<>();
        
        zipItems.stream()
                .map(ZipFileItem::cloneItem)
                .forEach(clonedItems::add);
        
        return new ZIPDocumentFile(clonedItems);
    }
    
    public void writeTo(OutputStream fileOut)
            throws IOException {
        
        final int BUFFER_SIZE = 32 * 1024;
        
        try (   BufferedOutputStream bufOut = new BufferedOutputStream(
                        fileOut, BUFFER_SIZE);
                ZipOutputStream zipOut = new ZipOutputStream(bufOut, UTF_8)) {
                            
            for (ZipFileItem fileItem : zipItems) {
                zipOut.putNextEntry(fileItem.zipEntry);
                
                if (null != fileItem.data) {
                    zipOut.write(fileItem.data, 0, fileItem.data.length);
                }
                
                zipOut.closeEntry();
            }
            
            bufOut.flush();
        }
    }
    
    /**
     * Prüft ob der Eintrag mit dem übergebenen Namen in der ZIP Datei
     * enthalten ist.
     * 
     * @param name  Name, z.B. {@code 'word/content.xml'}.
     * 
     * @return      {@code true}, wenn der Eintrag gefunden wurde und
     *              vorhanden ist
     */
    public boolean hasZipFileItem(String name) {
        Objects.requireNonNull(name, "name");
        return findItemByName0(name).isPresent();
    }
    
    /**
     * Sucht alle Dateieinträge raus deren Dateiname mit dem übergebenen
     * übereinstimmt oder beginnt.
     * 
     * <p>Es erfolgt eine Unterscheidung zwischen Groß- und Kleinschreibung.</p>
     * 
     * @param nameStartsWith    Pfad/Name als Prefix.
     * 
     * @return                  Alle gefundenen Einträge die mit dem übergebenen
     *                          Prefix/Pfad beginnen.
     */
    public List<String> findItemsStartingWith(String nameStartsWith) {
        return zipItems.stream()
                .map(zipItem -> zipItem.zipEntry.getName())
                .filter(name -> name.startsWith(nameStartsWith))
                .collect(toList());
    }
    
    private Optional<ZipFileItem> findItemByName0(String name) {
        return zipItems.stream()
                .filter(zipItem -> zipItem.zipEntry.getName().equals(name))
                .findFirst();
    }
    
    private ZipFileItem findItemByName(String name) {
        Optional<ZipFileItem> item = findItemByName0(name);
        
        if (item.isPresent() == false) {
            throw new InvalidDocumentFormatForImplementation(
                    formatMessage(MSG_MISSING_ITEM,
                            /* {0} */ name));
        }
        
        return item.get();
    }
    
    /**
     * Legt einen neuen Eintrag für den übergebenen Dateinamen in der ZIP Datei ab.
     * 
     * <p>Nach dem Anlegen kann mit {@link #overwrite(String, byte[])} entsprechend der Inhalt
     * geschrieben werden.</p>
     * 
     * @param fileName      Dateiname
     */
    public void createNewFileInZip(String fileName) {
        final ZipFileItem zipItem = new ZipFileItem();
        
        zipItem.zipEntry = new ZipEntry(fileName);
        
        zipItem.zipEntry.setSize(0);
        zipItem.zipEntry.setCompressedSize(-1L);
        
        zipItem.zipEntry.setCreationTime(FileTime.from(Instant.now()));
        zipItem.zipEntry.setLastAccessTime(FileTime.from(Instant.now()));
        zipItem.zipEntry.setLastModifiedTime(FileTime.from(Instant.now()));
        
        this.zipItems.add(zipItem);
    }
    
    public ZIPDocumentFile overwrite(String name, byte[] data) {
        final ZipFileItem zipItem = findItemByName(name);
        
        zipItem.zipEntry.setSize(data.length);
        zipItem.zipEntry.setCompressedSize(-1L);
        
        zipItem.zipEntry.setCreationTime(FileTime.from(Instant.now()));
        zipItem.zipEntry.setLastAccessTime(FileTime.from(Instant.now()));
        zipItem.zipEntry.setLastModifiedTime(FileTime.from(Instant.now()));
        
        zipItem.data = Arrays.copyOf(data, data.length);
        
        return this;
    }
    
    public byte[] read(String name) {
        final ZipFileItem zipItem = findItemByName(name);
        
        return Arrays.copyOf(zipItem.data, zipItem.data.length);
    }
    
    private void readZipFile(Path inZipFile) {
        try {
            readZipFile0(inZipFile);
            
        } catch (IOException exception) {
            throw new UncheckedIOException(exception);
        }
    }
    
    private void readZipFile0(Path inZipFile)
            throws IOException {
        
        try (ZipFile zipFile = new ZipFile(inZipFile.toFile(), UTF_8)) {
            zipFile.stream().forEach(entry -> readZipFileEntry(zipFile, entry));
        }
    }
    
    private void readZipFileEntry(ZipFile zipFile, ZipEntry zipEntry) {
        final ZipFileItem item = new ZipFileItem();
        item.zipEntry = zipEntry;
        item.zipEntry.setCompressedSize(-1L);
        
        zipItems.add(item);

        if (zipEntry.isDirectory()) {
            return;
        }
        
        final int size = (int) zipEntry.getSize() + 16;
        final ByteArrayOutputStream byteOut = new ByteArrayOutputStream(size);
        
        try (InputStream inZipEntry = zipFile.getInputStream(zipEntry)) {
            copy(inZipEntry, byteOut);    
            
        } catch (IOException ioException) {
            throw new UncheckedIOException(ioException);
        }
        
        item.data = byteOut.toByteArray();
    }
    
    private static class ZipFileItem {
        private ZipEntry zipEntry;
        private byte[] data;
        
        public ZipFileItem cloneItem() {
            ZipFileItem newItem = new ZipFileItem();
            newItem.zipEntry = new ZipEntry(zipEntry);
            
            if (null != data) {
                newItem.data = Arrays.copyOf(data, data.length);
            }
            
            return newItem;
        }
    }
    
    ////////////////////////////////////////////////////////////////////////////
    
    private static int copy(InputStream input, OutputStream output)
            throws IOException {
        
        long count = copyLarge(input, output);
        
        if (count > Integer.MAX_VALUE) {
            return -1;
        }
        
        return (int) count;
    }

    private static long copyLarge(InputStream input, OutputStream output)
            throws IOException {
        
        return copyLarge(input, output, new byte[16 * 1024]);
    }

    private static long copyLarge(InputStream input, OutputStream output,
            byte[] buffer) throws IOException {
        
        long count = 0;
        int n = 0;
        while (-1 != (n = input.read(buffer))) {
            output.write(buffer, 0, n);
            count += n;
        }
        return count;
    }
    
}
