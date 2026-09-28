package com.timematters.qstar.configuration;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

public enum StageType {
    STAGE_DEVELOP("develop"),
    STAGE_LOCAL("local"),
    STAGE_INTEGRATION("integration"),
    STAGE_PRODUCTION("production");

    private String value;

    StageType(String value) {
        this.value = value;
    }

    @JsonValue
    public String getValue() {
        return value;
    }

    @Override
    public String toString() {
        return String.valueOf(value);
    }

    @JsonCreator
    public static StageType fromValue(String value) {
        for (StageType stageType : StageType.values()) {
            if (stageType.value.equals(value)) {
                return stageType;
            }
        }
        throw new IllegalArgumentException("Unexpected value '" + value + "'");
    }
}
