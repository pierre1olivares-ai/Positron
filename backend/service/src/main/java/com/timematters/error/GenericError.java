package com.timematters.error;

import com.remondis.remap.Mapper;
import com.remondis.remap.Mapping;
import com.timematters.qstar.api.model.ErrorATO;
import java.util.Objects;

public class GenericError extends Throwable {
    private static final long serialVersionUID = 6580015483223535904L;

    private String id;
    private Integer status = 500;
    private String code;
    private String title;
    private String details;

    public GenericError id(String id) {
        this.id = id;
        return this;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public GenericError status(Integer status) {
        this.status = status;
        return this;
    }

    public Integer getStatus() {
        return status;
    }

    public void setStatus(Integer status) {
        this.status = status;
    }

    public GenericError code(String code) {
        this.code = code;
        return this;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public GenericError title(String title) {
        this.title = title;
        return this;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public GenericError details(String details) {
        this.details = details;
        return this;
    }

    public String getDetails() {
        return details;
    }

    public void setDetails(String details) {
        this.details = details;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        GenericError errorResponse = (GenericError) o;
        return Objects.equals(this.id, errorResponse.id)
                && Objects.equals(this.status, errorResponse.status)
                && Objects.equals(this.code, errorResponse.code)
                && Objects.equals(this.title, errorResponse.title)
                && Objects.equals(this.details, errorResponse.details);
    }

    @Override
    public int hashCode() {
        return Objects.hash(id, status, code, title, details);
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("class Error {\n");
        sb.append("    id: ").append(toIndentedString(id)).append("\n");
        sb.append("    status: ").append(toIndentedString(status)).append("\n");
        sb.append("    code: ").append(toIndentedString(code)).append("\n");
        sb.append("    title: ").append(toIndentedString(title)).append("\n");
        sb.append("    details: ").append(toIndentedString(details)).append("\n");
        sb.append("}");
        return sb.toString();
    }

    private String toIndentedString(Object o) {
        if (o == null) {
            return "null";
        }
        return o.toString().replace("\n", "\n    ");
    }

    public ErrorATO toErrorResponse() {
        Mapper<GenericError, ErrorATO> mapper =
                Mapping.from(GenericError.class)
                        .to(ErrorATO.class)
                        .omitOtherSourceProperties()
                        .omitInDestination(ErrorATO::getMeta)
                        .mapper();

        return mapper.map(this);
    }
}
