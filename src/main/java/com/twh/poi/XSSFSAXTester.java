package com.twh.poi;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.*;

import static com.twh.poi.MapContentsHandler.DONE_MAP;

public class XSSFSAXTester {
    public static void main(String[] args) {
        try {
            InputStream inputStream = XSSFSAXTester.class.getClassLoader().getResourceAsStream("abc.xlsx");
            OPCPackage pkg = OPCPackage.open(inputStream);
            XSSFReader r = new XSSFReader(pkg);
            BlockingQueue<Map<String, String>> blockingQueue = new ArrayBlockingQueue<>(200);
            XMLReader parser = fetchSheetParser(blockingQueue, pkg);
            InputStream sheet = r.getSheetsData().next();
            InputSource sheetSource = new InputSource(sheet);
            ExecutorService executor = new ThreadPoolExecutor(10, 200,
                    10L, TimeUnit.SECONDS,
                    new ArrayBlockingQueue<>(1000));
            new Thread(() -> {
                while (true) {
                    try {
                        Map data = blockingQueue.take();
                        if (data == DONE_MAP) {
                            break;
                        }
                        execute(executor,data);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                }
            }).start();
            parser.parse(sheetSource);
            executor.shutdown();
            executor.awaitTermination(60L, TimeUnit.SECONDS);
            sheet.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (OpenXML4JException e) {
            e.printStackTrace();
        } catch (SAXException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    public static void execute(ExecutorService executor, Map data) {
        try {
            executor.execute(() -> {
                try {
                    TimeUnit.MILLISECONDS.sleep(100L);
                    System.out.println(Thread.currentThread().getName() + ":" + "执行完毕" + data);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            });
        } catch (RejectedExecutionException e) {
            try {
                TimeUnit.SECONDS.sleep(1);
                execute(executor, data);
            } catch (InterruptedException e1) {
                e1.printStackTrace();
            }
        }
    }

    public static XMLReader fetchSheetParser(BlockingQueue<Map<String, String>> blockingQueue, OPCPackage pkg) {
        try {
            XMLReader parser = SAXHelper.newXMLReader();
            DataFormatter formatter = new DataFormatter();
            XSSFReader r = new XSSFReader(pkg);

            //        ContentHandler handler = new LineToMapHandler(blockingQueue, sst);
            ContentHandler handler = new XSSFSheetXMLHandler(r.getStylesTable(),
                    new ReadOnlySharedStringsTable(pkg),
                    new MapContentsHandler(blockingQueue), formatter,
                    false);
            parser.setContentHandler(handler);

            return parser;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
