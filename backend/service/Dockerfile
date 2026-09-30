# This is based on the examples given at
# https://github.com/bmuschko/gradle-docker-plugin/wiki/examples

FROM tmregistry.azurecr.io/base_java_21_jdk_ubuntu:latest
# jar file has intentionally stripped off version suffixes

ENV JAVA_APP_JAR=qstar.jar
ENV AB_OFF=true

EXPOSE 8080

COPY ./build/libs/qstar-*.jar /app/qstar.jar
# Telemetry is opt-in. The checked-in template agent configuration is not activated.
# Leave room for metaspace/native allocations inside the 750Mi container limit.
CMD ["java", "-XX:MaxRAMPercentage=65.0", "-jar", "/app/qstar.jar"]
