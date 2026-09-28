# Custom OpenAPI codegen templates — fixes applied

These override openapi-generator's built-in "spring" library templates (`templateDir` in
`build.gradle`). They came from the template zip already customized, but hadn't been kept in
sync with the Spring Boot 3.3.4 / Jakarta EE upgrade the rest of build.gradle targets. Fixed here
so `./gradlew compileJava` actually succeeds against this project's declared dependency versions:

- **`model.mustache`**: hardcoded `import com.timematters.globaloffer.model.AbstractModel;` —
  "globaloffer" is a real internal project name this template was evidently extracted from
  (same evidence trail as `kubernetes.yaml`'s leftover "globaloffer" port name and configmap
  entries — see the repo-root fix notes). `AbstractModel` itself is real and load-bearing:
  `pojo.mustache` extends every parent-less generated model with it
  (`{{^parent}} extends AbstractModel {{/parent}}`), but no such class shipped in the template
  zip for *any* project. Since the import already hardcoded the source project's own package
  rather than solving this generically, the working assumption is each project using this
  template provides its own — done here as
  `src/main/java/com/timematters/qstar/model/AbstractModel.java` (empty; nothing calls a method
  on it, it's purely a common supertype), with the import corrected to match. Reusing this
  template for another project means updating both. Also hardcoded `javax.validation.*` /
  `javax.xml.bind.*` — renamed to `jakarta.validation.*` / `jakarta.xml.bind.*`, required for
  bean validation annotations to resolve under Spring Boot 3.
- **`api.mustache`**: same `javax.validation.*` → `jakarta.validation.*` fix.
- **`apiController.mustache`**: same `javax.validation.*` → `jakarta.validation.*` fix, for
  completeness — not currently exercised, since `interfaceOnly: true` (added to build.gradle's
  `openApiGenerate` config, also missing from the template as shipped) skips generating this
  file. Without `interfaceOnly`, it produces a concrete `@Controller` with every method stubbed
  to `NOT_IMPLEMENTED`, which collides with a hand-written controller of the same name/interface
  — that's why `SampleApiController.java` in the original template had to be a manually
  reconciled copy rather than something regenerated on every build.
