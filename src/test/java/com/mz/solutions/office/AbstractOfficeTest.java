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
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.junit.jupiter.api.BeforeAll;

public abstract class AbstractOfficeTest {

    protected static final Path TEST_SOURCE_DIRECTORY = Paths.get("src", "test", "resources", "com", "mz", "solutions", "office");
    protected static final Path TESTS_OUTPUT_PATH = Paths.get("target").resolve("test-resources");
    
    @BeforeAll
    public static void beforeTestRun() {
        try {
            Files.createDirectories(TESTS_OUTPUT_PATH);
        } catch (IOException e) {
            e.printStackTrace();
        }
        
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

        String fileName = inputPath.toFile().getName();

        int dotIndex = fileName.lastIndexOf('.');
        String extension = fileName.substring(dotIndex);
        String nameWithoutExtension = fileName.substring(0, dotIndex);

        if (suffix == null) {
            return TESTS_OUTPUT_PATH.resolve(nameWithoutExtension + "_Output" + extension);
        } else {
            return TESTS_OUTPUT_PATH.resolve(nameWithoutExtension + "_" + suffix + "_Output" + extension);
        }
    }

    protected final void processOpenDocument(DataPage page, Path docInput, Path docOutput) {


        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newOpenOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(docInput);

        document.generate(page, ResultFactory.toFile(docOutput));
    }

    protected final void processWordDocument(DataPage page, Path docInput, Path docOutput) {


        final OfficeDocumentFactory docFactory = OfficeDocumentFactory.newMicrosoftOfficeInstance();
        final OfficeDocument document = docFactory.openDocument(docInput);

        document.generate(page, ResultFactory.toFile(docOutput));
    }

    protected static String randStr(int length) {
        final char[] ALPHA = "abcdefghijklmnopqrstuvwxyzABDEFGHIJKLMNOPQRSTUVWXYZ0123456789 -"
                .toCharArray();

        final StringBuilder buildResult = new StringBuilder(length);
        for (int i = 0; i < length; i++) {
            final int index = (int) (Math.random() * ALPHA.length);
            buildResult.append(ALPHA[index]);
        }

        return buildResult.toString();
    }

}
