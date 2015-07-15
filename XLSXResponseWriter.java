package org.stig.solr.response;

import org.apache.lucene.index.IndexableField;
import org.apache.solr.common.SolrDocument;
import org.apache.solr.common.SolrDocumentList;
import org.apache.solr.common.SolrException;
import org.apache.solr.common.params.SolrParams;
import org.apache.solr.common.util.DateUtil;
import org.apache.solr.common.util.NamedList;
import org.apache.solr.request.SolrQueryRequest;
import org.apache.solr.response.SolrQueryResponse;
import org.apache.solr.response.RawResponseWriter;
import org.apache.solr.response.TextResponseWriter;
import org.apache.solr.response.QueryResponseWriter;
import org.apache.solr.response.ResultContext;
import org.apache.solr.schema.FieldType;
import org.apache.solr.schema.SchemaField;
import org.apache.solr.schema.StrField;
import org.apache.solr.search.DocList;
import org.apache.solr.search.ReturnFields;
import org.apache.solr.core.SolrCore;
import org.apache.solr.common.util.ContentStream;
import org.apache.solr.common.util.ContentStreamBase;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.Writer;
import java.io.CharArrayWriter;
import java.io.StringWriter;
import java.io.PrintWriter;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PipedInputStream;
import java.io.PipedOutputStream;
import java.awt.Color;
import java.util.*;

/**
 *
 */

public class XLSXResponseWriter extends RawResponseWriter {

  @Override
  public void init(NamedList n) {
  }

  @Override
  public void write(OutputStream out, SolrQueryRequest req, SolrQueryResponse rsp) throws IOException {
    //throwaway arraywriter just to satisfy super requirements; we're grabbing
    //all writes before they go to it anyway
    XLSXWriter w = new XLSXWriter(new CharArrayWriter(), req, rsp);
    try {
      w.writeResponse(out);
    } finally {
      w.close();
    }
  }

  @Override
  public String getContentType(SolrQueryRequest request, SolrQueryResponse response) {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
}

class XLSXWriter extends TextResponseWriter {

  SolrQueryRequest req;
  SolrQueryResponse rsp;

  Logger log = LoggerFactory.getLogger(SolrCore.class);
  Calendar cal;  

  class SerialWriteWorkbook {
    SXSSFWorkbook swb; 
    Sheet sh; 

    XSSFCellStyle headerStyle;
    int rowIndex;
    Row curRow;
    int cellIndex;

    public SerialWriteWorkbook() {
        this.swb = new SXSSFWorkbook(100);
        this.sh = this.swb.createSheet();

        this.rowIndex = 0;

        this.headerStyle = (XSSFCellStyle)swb.createCellStyle();
        this.headerStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
        //solid fill
        this.headerStyle.setFillPattern((short)1);
        Font headerFont = swb.createFont();
        headerFont.setFontHeightInPoints((short)14);
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        this.headerStyle.setFont(headerFont);
    }

    public void addRow() {
        curRow = sh.createRow(rowIndex++);
        cellIndex = 0;
    }

    public void setHeaderRow() {
        curRow.setHeightInPoints((short)21);
    }
    
    //sets last created cell to have header style
    public void setHeaderCell() {
        curRow.getCell(cellIndex - 1).setCellStyle(this.headerStyle);
    }

    //set the width of the most recently created column
    public void setColWidth(int charWidth) {
        //width is set in units of 1/256th of a character width for some reason
        this.sh.setColumnWidth(cellIndex - 1, 256*charWidth);
    }

    public void writeCell(String value) {
        Cell cell = curRow.createCell(cellIndex++);
        cell.setCellValue(value);
    }

    public void flush(OutputStream out) {
        try {
            swb.write(out);
            log.info("Flush complete");
        } catch (IOException e) {
            StringWriter sw = new StringWriter();
            e.printStackTrace(new PrintWriter(sw));
            String stacktrace = sw.toString();
            log.info("Failed to export to XLSX - "+stacktrace);
        }finally {
            //clean temp files created by poi
            swb.dispose();
        }
    }
  }

  SerialWriteWorkbook wb = new SerialWriteWorkbook();

  static class Field {
    String name;
    SchemaField sf;

    // used to collect values
    List<IndexableField> values = new ArrayList<IndexableField>(1);  // low starting amount in case there are many fields
    int tmp;
  }

  Map<String,Field> xlFields = new LinkedHashMap<String,Field>();

  public XLSXWriter(Writer writer, SolrQueryRequest req, SolrQueryResponse rsp){
    super(writer, req, rsp);
    this.req = req;
    this.rsp = rsp;
  }

  public void writeResponse(OutputStream out) throws IOException {
    SolrParams params = req.getParams();
    log.info("Beginning export");

    Collection<String> fields = returnFields.getRequestedFieldNames();
    Object responseObj = rsp.getValues().get("response");
    boolean returnOnlyStored = false;
    if (fields==null||returnFields.hasPatternMatching()) {
      if (responseObj instanceof SolrDocumentList) {
        // get the list of fields from the SolrDocumentList
        if(fields==null) {
          fields = new LinkedHashSet<String>();
        }
        for (SolrDocument sdoc: (SolrDocumentList)responseObj) {
          fields.addAll(sdoc.getFieldNames());
        }
      } else {
        // get the list of fields from the index
        Collection<String> all = req.getSearcher().getFieldNames();
        if(fields==null) {
          fields = all;
        }
        else {
          fields.addAll(all);
        }
      }
      if (returnFields.wantsScore()) {
        fields.add("score");
      } else {
        fields.remove("score");
      }
      returnOnlyStored = true;
    }

    for (String field : fields) {
       if (!returnFields.wantsField(field)) {
         continue;
       }
      if (field.equals("score")) {
        Field xlField = new Field();
        xlField.name = "score";
        xlFields.put("score", xlField);
        continue;
      }

      SchemaField sf = schema.getFieldOrNull(field);
      if (sf == null) {
        FieldType ft = new StrField();
        sf = new SchemaField(field, ft);
      }
      
      // Return only stored fields, unless an explicit field list is specified
      if (returnOnlyStored && sf != null && !sf.stored()) {
        continue;
      }

      Field xlField = new Field();
      xlField.name = field;
      xlField.sf = sf;
      xlFields.put(field, xlField);
    }


    NiceMetadataNames niceMap = new NiceMetadataNames();

    wb.addRow();
    //write header
    for (Field xlField : xlFields.values()) {
        String printName = xlField.name;
        int colWidth = 14;

        NameAndWidth nextField = niceMap.get(xlField.name);
        if (nextField != null) {
            printName = nextField.getName();;
            colWidth = nextField.getWidth();
        }
       
        writeStr(xlField.name, printName, false);
        wb.setColWidth(colWidth);
        wb.setHeaderCell();
    }
    wb.setHeaderRow();
    wb.addRow();

    //write rows
    if (responseObj instanceof ResultContext ) {
      writeDocuments(null, (ResultContext)responseObj, returnFields );
    }
    else if (responseObj instanceof DocList) {
      ResultContext ctx = new ResultContext();
      ctx.docs =  (DocList)responseObj;
      writeDocuments(null, ctx, returnFields );
    } else if (responseObj instanceof SolrDocumentList) {
      writeSolrDocumentList(null, (SolrDocumentList)responseObj, returnFields );
    }
    log.info("Export complete; flushing document");
    //flush to outputstream 
    wb.flush(out);
    wb = null;

  }

  @Override
  public void close() throws IOException {
    super.close();
  }

  @Override
  public void writeNamedList(String name, NamedList val) throws IOException {
  }

  @Override
  public void writeStartDocumentList(String name, 
      long start, int size, long numFound, Float maxScore) throws IOException
  {
    // nothing
  }

  @Override
  public void writeEndDocumentList() throws IOException
  {
    // nothing
  }

  //NOTE: a document cannot currently contain another document
  List tmpList;
  @Override
  public void writeSolrDocument(String name, SolrDocument doc, ReturnFields returnFields, int idx ) throws IOException {
    if (tmpList == null) {
      tmpList = new ArrayList(1);
      tmpList.add(null);
    }

    for (Field xlField : xlFields.values()) {
      Object val = doc.getFieldValue(xlField.name);
      int nVals = val instanceof Collection ? ((Collection)val).size() : (val==null ? 0 : 1);
      if (nVals == 0) {
        writeNull(xlField.name);
        continue;
      }

      if ((xlField.sf != null && xlField.sf.multiValued()) || nVals > 1) {
        Collection values;
        // normalize to a collection
        if (val instanceof Collection) {
          values = (Collection)val;
        } else {
          tmpList.set(0, val);
          values = tmpList;
        }

        writeArray(xlField.name, values.iterator());

      } else {
        // normalize to first value
        if (val instanceof Collection) {
          Collection values = (Collection)val;
          val = values.iterator().next();
        }
        writeVal(xlField.name, val);
      }
    }
    wb.addRow();
  }

  @Override
  public void writeStr(String name, String val, boolean needsEscaping) throws IOException {
    wb.writeCell(val);
  }

  @Override
  public void writeMap(String name, Map val, boolean excludeOuter, boolean isFirstVal) throws IOException {
  }

  @Override
  public void writeArray(String name, Iterator val) throws IOException {
    StringBuffer output = new StringBuffer();
    while (val.hasNext()) {
        Object v = val.next();
        if (v instanceof IndexableField) {
            IndexableField f = (IndexableField)v;
            output.append(f.stringValue() + "; ");
        } else {
            output.append(v.toString() + "; ");
        }
    }
    if (output.length() > 0) {
        output.deleteCharAt(output.length()-1);
        output.deleteCharAt(output.length()-1);
    }
    writeStr(name, output.toString(), false);
  }

  @Override
  public void writeNull(String name) throws IOException {
    wb.writeCell("");
  }

  @Override
  public void writeInt(String name, String val) throws IOException {
    wb.writeCell(val);
  }

  @Override
  public void writeLong(String name, String val) throws IOException {
    wb.writeCell(val);
  }

  @Override
  public void writeBool(String name, String val) throws IOException {
    wb.writeCell(val);
  }

  @Override
  public void writeFloat(String name, String val) throws IOException {
    wb.writeCell(val);
  }

  @Override
  public void writeDouble(String name, String val) throws IOException {
    wb.writeCell(val);
  }

  @Override
  public void writeDate(String name, Date val) throws IOException {
    StringBuilder sb = new StringBuilder(25);
    cal = DateUtil.formatDate(val, cal, sb);
    String outputDate = sb.substring(0,10);
    writeDate(name, outputDate);
  }

  @Override
  public void writeDate(String name, String val) throws IOException {
    wb.writeCell(val);
  }
}
