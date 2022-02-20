package com.alibaba.easyexcel.course.base.utils;

/**
 * 定位用的工具
 *
 * @author Jiaju Zhuang
 */
public class PositionUtils {

    /**
     * 根据行号 列号生成定位
     * 例如： rowIndex =2  column=3
     * 则返回： C2
     *
     * @param rowIndex 所在行的位置
     * @param column   所在列的位置
     * @return 位置
     */
    public static String position(int rowIndex, int column) {
        StringBuilder position = new StringBuilder();
        while (column >= 0) {
            position.append((char)('A' + (column % 26)));
            column = (column / 26) - 1;
        }
        return position.reverse().append(rowIndex + 1).toString();

    }
}
