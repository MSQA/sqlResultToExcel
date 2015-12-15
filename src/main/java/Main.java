import java.io.*;
import java.util.Date;

public class Main {

    public static void main(String[] args) {
        System.out.println("Hello World!");
        //从桌面读取文件
        File inputFile = new File("/Users/lizhengdong/Desktop/sqlResult.txt");
        BufferedReader reader = null;
        DocUtil.PlainTable table = DocUtil.newPlainTable();
        table.addSheet("红包查询");
        try {
            reader = new BufferedReader(new FileReader(inputFile));
            String temp = null;
            String[] tempArray = null;
            String[][] citys = CityList.citys;
            int cityIndex = -1;//城市那列的下标
            while ((temp = reader.readLine()) != null) {
                temp = temp.replace(" ", "");
                if (temp.contains("+")) {
                    continue;
                }
                table.addRow();
                tempArray = temp.split("\\|");

                for (int i = 0; i < tempArray.length; i++) {
                    String cell = tempArray[i];
                    if (cell.contains("城市")) {
                        cityIndex = i;
                    }
                    if (i == cityIndex) {
                        //对城市替换
                        for (String[] tempCity : citys) {
                            if (cell.equals(tempCity[0])) {
                                cell = tempCity[1];
                            }
                        }
                    }
                    if (cell.length() != 0) {
                        table.addCell(cell.trim());
                        System.out.print(cell + "  ");
                    }
                }
                System.out.println();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        //输出excel
        File outputFile = new File("/Users/lizhengdong/Desktop/2015-11-21至2015-11-30未审核名单.xlsx");
        try {
            FileOutputStream outputStream = new FileOutputStream(outputFile);
            table.output(outputStream, new Date().toString());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
