package com.timematters.qstar.model.mapper;

import com.remondis.remap.Mapper;

public interface ATOMapper<MODEL, ATO> {
    public Mapper<MODEL, ATO> getMapperToATO();

    public Mapper<ATO, MODEL> getMapperFromATO();

    public MODEL fromATO(ATO ato);

    public ATO toATO(MODEL model);
}
