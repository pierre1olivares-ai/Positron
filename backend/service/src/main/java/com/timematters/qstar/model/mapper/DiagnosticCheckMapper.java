package com.timematters.qstar.model.mapper;

import com.timematters.qstar.api.model.DiagnosticCheckATO;
import com.timematters.qstar.model.DiagnosticCheck;

public class DiagnosticCheckMapper extends SimpleATOMapper<DiagnosticCheck, DiagnosticCheckATO> {
    public DiagnosticCheckMapper() {
        super(DiagnosticCheck.class, DiagnosticCheckATO.class);
    }
}
