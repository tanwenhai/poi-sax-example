package com.twh.poi;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.*;
import java.util.concurrent.BlockingQueue;

import static com.twh.poi.MapContentsHandler.columnIndexFormString;

public class LineToMapHandler extends DefaultHandler {
    // 阻塞队列
    private final BlockingQueue<Map> blockingQueue;
    // 共享字符串表
    private final SharedStringsTable sst;
    // 当前处理的单元格行索引
    private int lineNum;
    // 当前处理的单元格列索引
    private int colum;
    // 上一行 换行之后将数字放到阻塞队列中
    private int prevLineNum = 0;
    private String lastContents;
    private boolean nextIsString;
    private boolean inlineStr;
    private Map lineMap = new HashMap(16);
    List<String> fields = new ArrayList<>(16);

    public LineToMapHandler(BlockingQueue<Map> blockingQueue, SharedStringsTable sst) {
        this.blockingQueue = blockingQueue;
        this.sst = sst;
    }

    public void parseLine(String c) {
        lineNum = Integer.parseInt(c.replaceAll("[A-Z]", ""));
        colum = columnIndexFormString(c.replaceAll(String.valueOf(lineNum), ""));
    }

    @Override
    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {
        // c => cell
        if(name.equals("c")) {
            // Print the cell reference
            // Figure out if the value is an index in the SST
            String cellType = attributes.getValue("t");
            // 单元格
            String r = attributes.getValue("r");
            System.out.println(r);
            parseLine(r);
            nextIsString = cellType != null && cellType.equals("s");
            inlineStr = cellType != null && cellType.equals("inlineStr");
        } else if (name.equals("dimension")) {
            attributes.getIndex("ref");
        }
        // Clear contents cache
        lastContents = "";
    }

    @Override
    public void endElement(String uri, String localName, String name)
            throws SAXException {
        // Process the last contents as required.
        // Do now, as characters() may be called more than once
        if(nextIsString) {
            Integer idx = Integer.valueOf(lastContents);
            lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            nextIsString = false;
        }

        // v => contents of a cell
        // Output after we've seen the string contents
        if(name.equals("v") || (inlineStr && name.equals("c"))) {
//            System.out.printf("%s ", lastContents);
            if (lineNum == 1) {
                fields.add(lastContents);
            } else {
                if (lineNum != prevLineNum) {
                    System.out.println();
                    if (prevLineNum > 1) {
                        try {
                            blockingQueue.put(lineMap);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }
                        lineMap = new HashMap();
                    }
                    prevLineNum = lineNum;
                }
                lineMap.put(fields.get(colum), lastContents);
            }
        }
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
        lastContents += new String(ch, start, length);
    }
}
