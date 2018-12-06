/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2016,   Moritz Riebe     (moritz.riebe@mz-solutions.de)
 *                       Andreas Zaschka  (andreas.zaschka@mz-solutions.de)
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

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;

abstract class StandardImageResource implements ImageResource {

    static final int BUFFER_WRITE_SIZE = 16 * 1024;
    static final int BUFFER_READ_SIZE = BUFFER_WRITE_SIZE * 4;
    private final ImageResourceType formatType;

    ////////////////////////////////////////////////////////////////////////////////////////////////
    private volatile byte[] imageDataCache;

    public StandardImageResource(ImageResourceType formatType) {
        this.formatType = formatType;
    }

    private static byte[] toByteArray(InputStream imageStream) {
        try (final ByteArrayOutputStream outByte = new ByteArrayOutputStream(BUFFER_WRITE_SIZE)) {
            byte[] buffer = new byte[BUFFER_READ_SIZE];
            int read = 0;

            while ((read = imageStream.read(buffer)) >= 0) {
                outByte.write(buffer, 0, read);
            }

            outByte.flush();

            return outByte.toByteArray();
        } catch (IOException ioException) {
            throw new UncheckedIOException(ioException);
        }
    }

    protected abstract byte[] loadImageDataNow();

    @Override
    public ImageResourceType getImageFormatType() {
        return formatType;
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////

    @Override
    public byte[] loadImageData() {
        if (null == imageDataCache) {
            this.imageDataCache = loadImageDataNow();
        }

        return this.imageDataCache;
    }

    static class LazyEmbedImgResFile extends StandardImageResource {

        protected final Path imageFile;

        public LazyEmbedImgResFile(Path imageFile, ImageResourceType formatType) {
            super(formatType);
            this.imageFile = imageFile;
        }

        @Override
        protected byte[] loadImageDataNow() {
            try {
                return Files.readAllBytes(imageFile);
            } catch (IOException ioException) {
                throw new UncheckedIOException(ioException);
            }
        }

    }

    static class EagerEmbedImgResFile extends LazyEmbedImgResFile {

        public EagerEmbedImgResFile(Path imageFile, ImageResourceType formatType) {
            super(imageFile, formatType);
            loadImageDataNow();
        }

    }

    static class EagerEmbedImgResData extends StandardImageResource {

        private final byte[] imageData;

        public EagerEmbedImgResData(InputStream imageStream, ImageResourceType formatType) {
            this(toByteArray(imageStream), formatType);
        }

        public EagerEmbedImgResData(byte[] imageData, ImageResourceType formatType) {
            super(formatType);
            this.imageData = imageData;
        }

        @Override
        protected byte[] loadImageDataNow() {
            return imageData;
        }

    }

    static class LocalImageResourceImpl extends LazyEmbedImgResFile implements LocalImageResource {

        public LocalImageResourceImpl(Path imageFile, ImageResourceType formatType) {
            super(imageFile, formatType);
        }

        @Override
        public Path getLocalResource() {
            return imageFile;
        }

    }

    static class ExternalImgResImpl extends StandardImageResource implements ExternalImageResource {

        private final URL imageURL;

        public ExternalImgResImpl(URL imageURL, ImageResourceType formatType) {
            super(formatType);
            this.imageURL = imageURL;
        }

        @Override
        protected byte[] loadImageDataNow() {
            try (InputStream imageStream = imageURL.openStream()) {
                return toByteArray(imageStream);
            } catch (IOException ioException) {
                throw new UncheckedIOException(ioException);
            }
        }

        @Override
        public URL getResourceURL() {
            return this.imageURL;
        }

    }


}
