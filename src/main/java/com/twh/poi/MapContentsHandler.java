package com.twh.poi;


import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.*;
import java.util.concurrent.BlockingQueue;

public class MapContentsHandler implements SheetContentsHandler {
    public static final Map<String, String> DONE_MAP = Collections.emptyMap();
    private final BlockingQueue<Map<String, String>> blockingQueue;
    private Map<String, String> lineMap;
    private final List<String> fields = new ArrayList<>(16);
    // 单元格列转数值映射MAP
    private static final HashMap<String, Integer> columnIndexCache = new HashMap<>();
    // 当前行
    private int currentLine;
    // 当前列
    private int currentColumn;

    private boolean isHeader;

    public MapContentsHandler(BlockingQueue<Map<String, String>> blockingQueue) {
        this.blockingQueue = blockingQueue;
    }

    /**
     * 单元格转数字
     * A:1 Z:26 AA:27
     * @param column
     * @return
     * @throws Exception
     */
    public static int columnIndexFormString(String column) throws IllegalArgumentException {
        if (columnIndexCache.containsKey(columnIndexCache)) {
            return columnIndexCache.get(column);
        }

        int index;
        char[] pStrings = column.toCharArray();
        if (pStrings.length == 1) {
            index = pStrings[0] - 'A';
        } else if (pStrings.length == 2) {
            index = (pStrings[0] - 'A' + 1) * 26 + pStrings[1] - 'A';
        } else if (pStrings.length == 3) {
            index = ((pStrings[0] - 'A') * 676) + ((pStrings[1] - 'A') * 26) + pStrings[2] - 'A';
        } else {
            throw new IllegalArgumentException("不支持");
        }

        return index;
    }

    @Override
    public void startRow(int rowNum) {
        currentLine = rowNum + 1;
        isHeader = currentLine == 1;
        int initialCapacity = 16;
        if (Objects.nonNull(lineMap)) {
            initialCapacity = lineMap.size();
        }
        lineMap = new HashMap<>(initialCapacity);
    }

    @Override
    public void endRow(int rowNum) {
        if (currentLine < 2) {
            return;
        }

        try {
            blockingQueue.put(lineMap);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        currentColumn = columnIndexFormString(cellReference.replaceAll(String.valueOf(currentLine),""));
        if (isHeader) {
            fields.add(formattedValue);
        } else {
            lineMap.put(fields.get(currentColumn), formattedValue);
        }
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        if (!isHeader) {
            try {
                blockingQueue.put(DONE_MAP);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }
}
