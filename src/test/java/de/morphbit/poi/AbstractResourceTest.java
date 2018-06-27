package de.morphbit.poi;

import java.io.File;
import java.net.URL;

public abstract class AbstractResourceTest {

    protected File getResourceFile(String filePath) {
        ClassLoader classLoader = getClass().getClassLoader();
        URL url = classLoader.getResource(filePath);
        return new File(url.getFile());
    }
}
