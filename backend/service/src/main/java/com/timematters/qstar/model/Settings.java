package com.timematters.qstar.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Data;

@Data
@JsonIgnoreProperties(ignoreUnknown = true)
public class Settings {
    private String msFormUrl = "";
    private String flowId = "";
}
