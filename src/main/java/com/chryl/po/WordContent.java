package com.chryl.po;

import lombok.Data;

/**
 * Created by Chr.yl on 2025/11/7.
 *
 * @author Chr.yl
 */
@Data
public class WordContent {
    private String level1; // 一级标题
    private String level2; // 二级标题
    private String level3; // 三级标题
    private String level4; // 四级标题
    private String level5; // 五级标题
    private String content; // 正文
}
