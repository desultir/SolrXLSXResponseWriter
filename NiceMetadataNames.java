/**
 * Define a translation of SOLR metadata field names to Nice names for 
 * printing in XLSX headers
 */

package org.stig.solr.response;
import java.util.HashMap;

public class NiceMetadataNames extends HashMap<String, NameAndWidth> {

    public NiceMetadataNames() {
        super();
        this.put("meta_type_1", new NameAndWidth("Nice Name 1", 14));
        this.put("long_meta_2", new NameAndWidth("Long Metadata name is long", 128));
    }
}

class NameAndWidth {
    private int width;
    private String name;

    public NameAndWidth(String name, int width) {
        this.name = name;
        this.width = width;
    }

    public String getName() {
        return this.name;
    }

    public int getWidth() {
        return this.width;
    }
}


