package com.timematters.qstar.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ProgressLogEntry {
    private Long id;
    private String ts;
    private String author;
    private Long authorId;
    private String authorEmail;
    private String text;
    private String saveWarning;
}
