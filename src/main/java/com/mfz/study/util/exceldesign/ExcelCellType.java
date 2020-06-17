package com.mfz.study.util.exceldesign;


public enum ExcelCellType {

    STRING(1),

    DECIMAL(2),

    _NONE(3),

    TIME(4),

    DATE(5);

    private final int code;

    ExcelCellType(int code) {
        this.code = code;
    }

    public static ExcelCellType getByCode(int code) {
        for (ExcelCellType type : values()) {
            if (type.code == code) {
                return type;
            }
        }
        throw new IllegalArgumentException("Invalid ExcelCellType code: " + code);
    }

    public int getCode() {
        return code;
    }

}
