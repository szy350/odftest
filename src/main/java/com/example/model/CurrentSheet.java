package com.example.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Sheet;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class CurrentSheet {

    private Sheet sheet;

    // row由sheet创建,与sheet绑定
    private int rowIndex;
}
