package com.pangchun.poi.write;

import com.pangchun.poi.support.bean.CommentBean;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * @author pangchun
 * @since 2021/6/15
 * @description 通用Excel写入类，提供通用写入方法
 */
public class CommonWrite {

    /**
     * 插入批注
     * @param workbook 工作簿
     * @param sheet 工作表
     * @param list 批注类的集合
     */
    public void insertComment(Workbook workbook, Sheet sheet, List<CommentBean> list) {
        // 获取创建工具
        CreationHelper factory = workbook.getCreationHelper();
        // 插入批注
        list.forEach(e -> {
            int firstRow = e.getFirstRow();
            int firstCol = e.getFirstCol();
            // 设置定位
            ClientAnchor anchor = factory.createClientAnchor();
            // 这里设置+1是为了使excel中批注更好地显示
            anchor.setRow1(firstRow);
            anchor.setCol1(firstCol);
            anchor.setRow2(e.getLastRow() + 1);
            anchor.setCol2(e.getLastCol() + 1);
            // 创建批注
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            Comment comment = drawing.createCellComment(anchor);
            RichTextString message = factory.createRichTextString(e.getMessage());
            comment.setString(message);
            // 添加批注
            Cell cell = sheet.getRow(firstRow).getCell(firstCol);
            if (cell.getCellComment() == null) {
                cell.removeCellComment();
                cell.setCellComment(comment);
            }
        });
    }
}
