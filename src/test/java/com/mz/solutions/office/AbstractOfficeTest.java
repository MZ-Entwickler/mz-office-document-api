/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2018,   Moritz Riebe     (moritz.riebe@mz-solutions.de),
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
package com.mz.solutions.office;

import com.mz.solutions.office.model.DataPage;
import com.mz.solutions.office.result.ResultFactory;
import org.junit.jupiter.api.AfterAll;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public abstract class AbstractOfficeTest {

    protected static final Path TESTS_OUTPUT_PATH = Paths.get("target").resolve("test-templates");

    @AfterAll
    public static void setup() {
        try {
            Files.createDirectories(TESTS_OUTPUT_PATH);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    protected Path outputPathOf(Path inputPath) {

        return outputPathOf(inputPath, null);
    }

    protected Path outputPathOf(Path inputPath, String suffix) {

        Path name = inputPath.getFileName();

        // null for empty paths and root-only paths
        if (name == null) {
            return inputPath;
        }

        String fileName = name.toString();
        int dotIndex = fileName.lastIndexOf('.');
        String extension = fileName.substring(dotIndex);
        String nameWithoutExtension = fileName.substring(0, dotIndex);

        if (dotIndex == -1) {
            return name;
        }

        if (suffix == null) {
            return inputPath.resolveSibling(nameWithoutExtension + "_Output" + extension);
        } else {
            return inputPath.resolveSibling(nameWithoutExtension + "_Output" + "_" + suffix + extension);
        }
    }

    protected final void processOpenDocument(DataPage page, Path docInput, String docOutput) {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(docInput);

        document.generate(page, ResultFactory.toFile(TESTS_OUTPUT_PATH.resolve(docOutput)));
    }

    protected final void processWordDocument(DataPage page, Path docInput, String docOutput) {
        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(docInput);

        document.generate(page, ResultFactory.toFile(TESTS_OUTPUT_PATH.resolve(docOutput)));
    }

}
