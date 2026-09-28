package com.timematters.qstar.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class DiagnosticCheck {
    private String name;
    private String status; // pass | warn | fail
    private String message;
}
