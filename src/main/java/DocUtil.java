import java.io.*;
import java.lang.ref.WeakReference;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author kailong.li
 */
public class DocUtil {

    // data management

    private static Map<String, WeakReference<byte[]>> dataCache = new HashMap<String, WeakReference<byte[]>>();




    // workbook

    /**
     * 16K
     */
    public static final int COL_LIMIT = SpreadsheetVersion.EXCEL2007.getMaxColumns();

    /**
     * 1M
     */
    public static final int ROW_LIMIT = SpreadsheetVersion.EXCEL2007.getMaxRows();

    private static final String ROW_OVERFLOW_SHEET_SUF = "-NEXT";

    public static PlainTable newPlainTable() {
        return new PlainTable();
    }

    public static StreamingPlainTable newStreamingPlainTable() {
        return new StreamingPlainTable();
    }

    /**
     * <pre>
     * DocUtil.PlainTable table = DocUtil.newPlainTable();
     * table.addSheet(&quot;2012-10&quot;);
     * table.addRow();
     * table.addCell(&quot;ID&quot;);
     * table.addCell(&quot;VALUE&quot;);
     * for (int i = 0; i &lt; 10; i ++) {
     *     table.addRow();
     *     table.addCell(i);
     *     table.addCell(&quot;value-&quot; + i);
     * }
     * table.output(new FileOutputStream(&quot;d:\\temp\\1.xls&quot;), &quot;DocUtil.PlainTable&quot;);
     * </pre>
     *
     * @author kailong.li
     */
    public static class PlainTable {

        protected XSSFWorkbook workbook;

        protected CellStyle dateCellStyle;

        protected CellStyle datetimeCellStyle;

        protected CellStyle timeCellStyle;

        protected Sheet sheet;

        protected int rowIndex = 0;

        protected Row row;

        protected int colIndex = 0;

        private PlainTable() {
            workbook = new XSSFWorkbook();
            DataFormat dataFormat = workbook.createDataFormat();

            dateCellStyle = workbook.createCellStyle();
            short dateFormatIndex = dataFormat.getFormat("yyyy-m-d");
            dateCellStyle.setDataFormat(dateFormatIndex);

            datetimeCellStyle = workbook.createCellStyle();
            short datetimeFormatIndex = dataFormat.getFormat("yyyy-m-d h:mm:ss");
            datetimeCellStyle.setDataFormat(datetimeFormatIndex);

            timeCellStyle = workbook.createCellStyle();
            short timeFormatIndex = dataFormat.getFormat("h:mm:ss");
            timeCellStyle.setDataFormat(timeFormatIndex);
        }

        public void addSheet(String name) {
            workbookCreateSheet(name);
            resetRowIndex();
            resetColIndex();
        }

        public void addRow() {
            if (rowIndex == ROW_LIMIT) {
                addSheet(sheet.getSheetName() + ROW_OVERFLOW_SHEET_SUF);
                resetRowIndex();
            }
            row = sheet.createRow(rowIndex++);
            resetColIndex();
        }

        public void addCell() {
            row.createCell(colIndex++);
        }

        public void addCell(Number number) {
            Cell cell = row.createCell(colIndex++);
            if (number != null) {
                // double
                cell.setCellValue(number.doubleValue());
            }
        }

        public void addCell(Date date, boolean datetimeFormat) {
            Cell cell = row.createCell(colIndex++);
            if (date != null) {
                // date
                cell.setCellValue(date);
                cell.setCellStyle(datetimeFormat ? datetimeCellStyle : dateCellStyle);
            }
        }

        public void addCell(Date date, int dateStyle) {
            Cell cell = row.createCell(colIndex++);
            if (date != null) {
                // date
                cell.setCellValue(date);
                switch (dateStyle) {
                    case 1:
                        cell.setCellStyle(dateCellStyle);
                        break;
                    case 2:
                        cell.setCellStyle(timeCellStyle);
                        break;
                    default:
                        break;
                }
            }
        }

        public void addCell(Object value) {
            Cell cell = row.createCell(colIndex++);
            if (value != null) {
                // string
                cell.setCellValue(value.toString());
            }
        }


        /**
         * @param out     will be flushed and closed
         * @param comment
         * @throws IOException
         */
        public void output(OutputStream out, String comment) throws IOException {
            workbook.getProperties().getCoreProperties().setDescription(comment);
            workbook.write(out);
            out.flush();
            out.close();
        }

        protected void workbookCreateSheet(String name) {
            sheet = workbook.createSheet(name);
        }

        protected void resetRowIndex() {
            rowIndex = 0;
        }

        protected void resetColIndex() {
            colIndex = 0;
        }
    }

    /**
     * <pre>
     * DocUtil.StreamingPlainTable table = DocUtil.newStreamingPlainTable();
     * table.addSheet(&quot;2012-10&quot;);
     * table.addRow();
     * table.addCell(&quot;ID&quot;);
     * table.addCell(&quot;VALUE&quot;);
     * boolean stop = false;
     * for (int i = 0; i &lt; 10; i ++) {
     *     table.addRow();
     *     table.addCell(i);
     *     table.addCell(&quot;value-&quot; + i);
     *     if (stop) {
     *         table.dispose();
     *         return;
     *     }
     * }
     * table.output(new FileOutputStream(&quot;d:\\temp\\1.xls&quot;), &quot;DocUtil.PlainTable&quot;);
     * </pre>
     *
     * @author kailong.li
     */
    public static class StreamingPlainTable extends PlainTable {

        protected SXSSFWorkbook streamingWorkbook;

        private StreamingPlainTable() {
            super();
            streamingWorkbook = new SXSSFWorkbook(workbook);
        }

        @Override
        protected void workbookCreateSheet(String name) {
            sheet = streamingWorkbook.createSheet(name);
        }

        /**
         * auto dispose
         *
         * @param out     will be flushed and closed
         * @param comment
         * @throws IOException
         * @see #dispose
         */
        @Override
        public void output(OutputStream out, String comment) throws IOException {
            workbook.getProperties().getCoreProperties().setDescription(comment);
            streamingWorkbook.write(out);
            out.flush();
            out.close();
            streamingWorkbook.dispose();
        }

        /**
         * delete all temp files
         */
        public void dispose() {
            streamingWorkbook.dispose();
        }
    }

    // TODO
    public static class CellRangeTable {

    }

    // TODO
    public static class StylingTable {

    }


    public static void main(String[] a) throws FileNotFoundException, IOException {
//        DocUtil.PlainTable table = DocUtil.newPlainTable();
        StreamingPlainTable table = DocUtil.newStreamingPlainTable();
        table.addSheet("2012-10");
        table.addRow();
        table.addCell("ID");
        table.addCell("VALUE");
        for (int i = 0; i < 10; i++) {
            table.addRow();
            table.addCell(i);
            table.addCell("value值" + i);
        }
        table.output(new FileOutputStream("d:\\temp\\1.xls"), "made by DocUtil");

    }


}
