# This is based on the examples given at
# https://github.com/bmuschko/gradle-docker-plugin/wiki/examples

FROM tmregistry.azurecr.io/base_java_21_jdk_ubuntu:latest
# jar file has intentionally stripped off version suffixes

ENV JAVA_APP_JAR qstar.jar
ENV AB_OFF true

EXPOSE 8080:8080

COPY ./build/libs/qstar-*.jar /app/qstar.jar
ADD ./newrelic/* /app/

CMD ["java", "-javaagent:/app/newrelic.jar", "-Dnewrelic.config.file=/app/newrelic.yml", "-Xmx1200m", "-jar", "/app/qstar.jar"]