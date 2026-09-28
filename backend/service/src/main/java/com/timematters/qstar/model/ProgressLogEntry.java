package com.timematters.qstar.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ProgressLogEntry {
    private String ts;
    private String author;
    private String text;
}
