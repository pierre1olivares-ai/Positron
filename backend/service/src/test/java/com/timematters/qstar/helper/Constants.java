package com.timematters.qstar.helper;

import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneOffset;
import java.util.UUID;

public class Constants {
    public static final String STATIC_UUID = UUID.randomUUID().toString();
    public static final OffsetDateTime DATE_TIME =
            OffsetDateTime.of(
                    LocalDateTime.of(2030, 05, 12, 05, 45), ZoneOffset.ofHoursMinutes(0, 0));
    ;
}
