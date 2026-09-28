package com.timematters.qstar.model;

// pojo.mustache has every generated *ATO extend this when the OpenAPI schema doesn't declare a
// parent — pojo.mustache: "{{^parent}} extends AbstractModel {{/parent}}". No such class shipped
// in the template zip; the import that referenced it (model.mustache) hardcoded the concrete
// package of whatever project the template was extracted from
// (com.timematters.globaloffer.model.AbstractModel), suggesting each project using this template
// is expected to provide its own. Left empty since nothing in this codebase calls a method on it
// — it exists purely so generated ATOs share a common supertype, same as the original.
public abstract class AbstractModel {}
