package com.timematters.qstar.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class AccessEntry {
    private String email;
    private String role; // admin | qm | owner | reader
}
