package com.timematters.qstar.model.mapper;

import com.remondis.remap.Mapper;
import com.remondis.remap.Mapping;
import com.timematters.qstar.api.model.IssueATO;
import com.timematters.qstar.model.Issue;

public class IssueMapper implements ATOMapper<Issue, IssueATO> {

    private final ProgressLogEntryMapper progressLogEntryMapper;

    public IssueMapper() {
        this.progressLogEntryMapper = new ProgressLogEntryMapper();
    }

    @Override
    public Mapper<Issue, IssueATO> getMapperToATO() {
        return Mapping.from(Issue.class)
                .to(IssueATO.class)
                .useMapper(progressLogEntryMapper.getMapperToATO())
                .mapper();
    }

    @Override
    public Mapper<IssueATO, Issue> getMapperFromATO() {
        return Mapping.from(IssueATO.class)
                .to(Issue.class)
                .useMapper(progressLogEntryMapper.getMapperFromATO())
                .mapper();
    }

    @Override
    public Issue fromATO(IssueATO ato) {
        return getMapperFromATO().map(ato);
    }

    @Override
    public IssueATO toATO(Issue model) {
        return getMapperToATO().map(model);
    }
}
