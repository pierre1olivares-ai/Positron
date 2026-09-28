package com.timematters.error;

import java.util.ArrayList;
import java.util.List;

public class ValidationViolations {
    private List<ValidationViolation> validationViolations = new ArrayList<>();

    public void add(String source, String error, Object value) {
        validationViolations.add(new ValidationViolation(source, error, String.valueOf(value)));
    }

    public void add(ValidationViolations validationViolations) {
        this.validationViolations.addAll(validationViolations.getValidationViolations());
    }

    public List<ValidationViolation> getValidationViolations() {
        return validationViolations;
    }

    public boolean isEmpty() {
        return this.validationViolations.isEmpty();
    }
}
