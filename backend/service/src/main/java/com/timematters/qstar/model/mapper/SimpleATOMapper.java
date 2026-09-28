package com.timematters.qstar.model.mapper;

import com.remondis.remap.Mapper;
import com.remondis.remap.Mapping;

public class SimpleATOMapper<MODEL, ATO> implements ATOMapper<MODEL, ATO> {

    private Mapper<MODEL, ATO> mapperToATO;
    private Mapper<ATO, MODEL> mapperFromATO;

    public SimpleATOMapper(Class<MODEL> model, Class<ATO> ato) {
        mapperToATO = Mapping.from(model).to(ato).mapper();
        mapperFromATO = Mapping.from(ato).to(model).mapper();
    }

    @Override
    public Mapper<MODEL, ATO> getMapperToATO() {
        return mapperToATO;
    }

    @Override
    public Mapper<ATO, MODEL> getMapperFromATO() {
        return mapperFromATO;
    }

    @Override
    public MODEL fromATO(ATO ato) {
        return mapperFromATO.map(ato);
    }

    @Override
    public ATO toATO(MODEL model) {
        return mapperToATO.map(model);
    }
}
