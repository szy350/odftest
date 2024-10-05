package com.example.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Row;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class CurrentRow {

    private Row currentRow;

    private int currentCellIndex;
}
