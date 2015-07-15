# SolrXLSXResponseWriter
XLSXResponseWriter
Solr XLSX ResponseWriter 

An XLSX Response Writer for Solr4.6.0. Should work with up to Solr4.x but hasn't been tested - porting to Solr5 is future work. If you want to test it and let me know that'd be great.

Dependencies:
* Apache POI 3.11 (poi.jar, poi-ooxml.jar and poi-ooxml-schemas.jar)
* Solr 4.6 jars (look in tomcat/webapps/solr/WEB-INF/lib/)
* SLF4j jars


Instructions:
* Edit the map in NiceMetadataNames.java to provide the XLSXResponseWriter a lookup for translating your ugly metadatanames into user friendly text strings. 

* Make sure all the above are on your classpath and make build

* Once you've built, pull the jar file out of ./build and drop it somewhere Solr will find it (ie /home/solr/collection1/lib/dist)

* Create a queryResponseHandler in solrconfig.xml:

    &lt;queryResponseWriter name="xlsx" class="org.stig.solr.response.XLSXResponseWriter" /&gt;

* Create a requesthandler which uses it:
    
    &lt;requestHandler name="/exportinv" class="solr.SearchHandler"&gt;

      &lt;lst name="defaults"&gt;
      
       &lt;str name="wt"&gt;xlsx&lt;/str&gt;
       
...
