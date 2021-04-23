package ru.curs.xylophone.descriptor;

import com.fasterxml.jackson.annotation.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.yaml.YAMLFactory;

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

    public void jsonSerialize(OutputStream json_stream) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.setSerializationInclusion(JsonInclude.Include.NON_NULL);
        mapper.writerWithDefaultPrettyPrinter().writeValue(json_stream, this);
    }

    //deserializes and json descriptors and yaml descriptors
    public static DescriptorElement yamlDeserialize(InputStream yaml_stream) throws IOException {
        ObjectMapper mapper = new ObjectMapper(new YAMLFactory());
        return mapper.readValue(yaml_stream, DescriptorElement.class);
    }
}
