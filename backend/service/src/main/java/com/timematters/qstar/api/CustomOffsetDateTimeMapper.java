package com.timematters.qstar.api;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import java.io.IOException;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import org.springframework.stereotype.Service;

@Service
public class CustomOffsetDateTimeMapper extends ObjectMapper {

    /** */
    private static final long serialVersionUID = -8508094245376588373L;

    public CustomOffsetDateTimeMapper() {
        super();
        this.registerModule(new JavaTimeModule());
        SimpleModule offsetDateTimeModule = new SimpleModule();
        offsetDateTimeModule.addSerializer(
                OffsetDateTime.class,
                new JsonSerializer<OffsetDateTime>() {
                    @Override
                    public void serialize(
                            OffsetDateTime offsetDateTime,
                            JsonGenerator jsonGenerator,
                            SerializerProvider serializerProvider)
                            throws IOException {
                        jsonGenerator.writeString(
                                DateTimeFormatter.ISO_OFFSET_DATE_TIME.format(offsetDateTime));
                    }
                });

        offsetDateTimeModule.addDeserializer(
                OffsetDateTime.class,
                new JsonDeserializer<OffsetDateTime>() {
                    @Override
                    public OffsetDateTime deserialize(
                            JsonParser jsonParser, DeserializationContext deserializationContext)
                            throws IOException {
                        String input = jsonParser.getText();
                        return OffsetDateTime.parse(input);
                    }
                });

        this.registerModule(offsetDateTimeModule);
    }
}
