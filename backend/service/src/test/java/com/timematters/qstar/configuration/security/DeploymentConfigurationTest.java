package com.timematters.qstar.configuration.security;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;
import org.yaml.snakeyaml.Yaml;

class DeploymentConfigurationTest {
    @SuppressWarnings("unchecked")
    @Test
    void ingressPreservesApiRouteAndTargetsDeclaredService() throws Exception {
        List<Map<String, Object>> documents = new ArrayList<>();
        try (var input = Files.newInputStream(Path.of("kubernetes.yaml"))) {
            new Yaml()
                    .loadAll(input)
                    .forEach(document -> documents.add((Map<String, Object>) document));
        }
        var ingress =
                documents.stream()
                        .filter(document -> "Ingress".equals(document.get("kind")))
                        .findFirst()
                        .orElseThrow();
        var service =
                documents.stream()
                        .filter(document -> "Service".equals(document.get("kind")))
                        .findFirst()
                        .orElseThrow();
        assertEquals("networking.k8s.io/v1", ingress.get("apiVersion"));
        var metadata = (Map<String, Object>) ingress.get("metadata");
        assertFalse(
                ((Map<?, ?>) metadata.get("annotations"))
                        .containsKey("nginx.ingress.kubernetes.io/rewrite-target"));
        var spec = (Map<String, Object>) ingress.get("spec");
        var rule = ((List<Map<String, Object>>) spec.get("rules")).get(0);
        var http = (Map<String, Object>) rule.get("http");
        var path = ((List<Map<String, Object>>) http.get("paths")).get(0);
        assertEquals("/api", path.get("path"));
        assertEquals("Prefix", path.get("pathType"));
        var target = (Map<String, Object>) ((Map<?, ?>) path.get("backend")).get("service");
        assertEquals(((Map<?, ?>) service.get("metadata")).get("name"), target.get("name"));
        assertEquals(8080, ((Map<?, ?>) target.get("port")).get("number"));
    }

    @SuppressWarnings("unchecked")
    @Test
    void pipelineUsesNestedServiceAndJava21() throws Exception {
        Map<String, Object> pipeline;
        try (var input = Files.newInputStream(Path.of("azure-pipelines.yml"))) {
            pipeline = new Yaml().load(input);
        }
        var steps = (List<Map<String, Object>>) pipeline.get("steps");
        var java =
                (Map<String, Object>)
                        steps.stream()
                                .filter(step -> "JavaToolInstaller@0".equals(step.get("task")))
                                .findFirst()
                                .orElseThrow()
                                .get("inputs");
        assertEquals("21", java.get("versionSpec"));
        var gradle =
                (Map<String, Object>)
                        steps.stream()
                                .filter(step -> "Gradle@4".equals(step.get("task")))
                                .findFirst()
                                .orElseThrow()
                                .get("inputs");
        assertEquals("backend/service", gradle.get("workingDirectory"));
        assertTrue(
                Files.isRegularFile(
                        Path.of("../..").resolve((String) gradle.get("gradleWrapperFile"))));
        assertEquals("Path", gradle.get("javaHomeOption"));
        assertEquals("$(JAVA_HOME)", gradle.get("jdkDirectory"));
        var docker =
                (Map<String, Object>)
                        steps.stream()
                                .filter(step -> "Docker@2".equals(step.get("task")))
                                .findFirst()
                                .orElseThrow()
                                .get("inputs");
        assertEquals("$(Build.SourcesDirectory)/backend/service", docker.get("buildContext"));
        assertEquals(
                "$(Build.SourcesDirectory)/backend/service/Dockerfile", docker.get("Dockerfile"));
    }
}
