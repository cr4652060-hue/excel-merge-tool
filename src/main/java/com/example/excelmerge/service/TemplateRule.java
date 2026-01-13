package com.example.excelmerge.service;

import java.util.List;

public record TemplateRule(
        String name,
        List<String> matchHeaders,
        List<String> requiredHeaders
) {
}