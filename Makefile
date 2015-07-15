SOLR_HOME = /home/solr
BUILD_DIR = build

clean : 
	rm -rf ./org
	cd ./$(BUILD_DIR) && rm -f XLSXResponseWriter.jar
     

build : ./$(BUILD_DIR)/lib/dist/XLSXResponseWriter.jar

./$(BUILD_DIR)/lib/dist/XLSXResponseWriter.jar : XLSXResponseWriter.java NiceMetadataNames.java
	javac  -d . ./NiceMetadataNames.java	
	javac  -d . ./XLSXResponseWriter.java	
	jar cf "XLSXResponseWriter.jar" ./org
	mv ./XLSXResponseWriter.jar ./$(BUILD_DIR)/XLSXResponseWriter.jar

FORCE:

