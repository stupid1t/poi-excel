package com.github.stupdit1t.excel.core.export;

import com.github.stupdit1t.excel.common.PoiCommon;
import com.github.stupdit1t.excel.core.AbsParent;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

public class OpsFooter<R> extends AbsParent<OpsSheet<R>> {

    /**
     * 复杂表头设计容器
     */
    List<ComplexCell> complexFooter;

    /**
     * 声明headers
     *
     * @param export parent
     */
    OpsFooter(OpsSheet<R> export) {
        super(export);
    }

    /**
     * 获取复杂表头设计
     *
     * @return List<ComplexHeader < R>>
     */
    public OpsFooter<R> text(String text, String location) {
        return this.textIndex(text, PoiCommon.coverRangeIndex(location));
    }

    /**
     * 获取复杂表头设计
     *
     * @return List<ComplexHeader < R>>
     */
    public OpsFooter<R> text(String text, String location, BiConsumer<Font, CellStyle> style) {
        this.textIndex(text, PoiCommon.coverRangeIndex(location), style);
        return this;
    }

    /**
     * 获取复杂表头设计
     *
     * @return List<ComplexHeader < R>>
     */
    public OpsFooter<R> textIndex(String text, Integer[] locationIndex) {
        return textIndex(text, locationIndex, null);
    }

    /**
     * 获取复杂表头设计
     *
     * @return List<ComplexHeader < R>>
     */
    public OpsFooter<R> textIndex(String text, Integer[] locationIndex, BiConsumer<Font, CellStyle> style) {
        if (complexFooter == null) {
            complexFooter = new ArrayList<>();
        }
        ComplexCell complexCell = new ComplexCell();
        complexCell.text = text;
        complexCell.locationIndex = locationIndex;
        complexCell.style = style;
        complexFooter.add(complexCell);
        return this;
    }
}
