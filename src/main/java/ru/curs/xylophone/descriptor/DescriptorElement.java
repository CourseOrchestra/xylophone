package ru.curs.xylophone.descriptor;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonGetter;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.LinkedList;
import java.util.List;

public final class DescriptorElement {
    private final String name;
    private final List<DescriptorOutputBase> sub_elements;

    public DescriptorElement(String name) {
        this.name = name;
        this.sub_elements = new LinkedList<>();
    }

    @JsonCreator(mode = JsonCreator.Mode.PROPERTIES)
    DescriptorElement(
            @JsonProperty("name") String name,
            @JsonProperty("output-steps") List<DescriptorOutputBase> sub_elements) {
        this.name = name;
        this.sub_elements = sub_elements;
    }

    @JsonGetter("name")
    public String getName() {
        return name;
    }

    @JsonGetter("output-steps")
    public List<DescriptorOutputBase> getSubElements() {
        return sub_elements;
    }

    public static DescriptorElement jsonDeserialize(InputStream json_stream) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        return mapper.readValue(json_stream, DescriptorElement.class);
    }

    public void jsonSerialize(OutputStream json_stream) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.setSerializationInclusion(JsonInclude.Include.NON_NULL);
        mapper.writerWithDefaultPrettyPrinter().writeValue(json_stream, this);
    }
}
