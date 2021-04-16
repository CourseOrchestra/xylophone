package ru.curs.xylophone.descriptor;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonGetter;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.util.LinkedList;
import java.util.List;

public final class DescriptorIteration extends DescriptorOutputBase {
    private final int index;
    private final int merge;
    private final Boolean horizontal;
    private final String regionName;
    private final List<DescriptorElement> elements;

    @JsonCreator(mode = JsonCreator.Mode.PROPERTIES)
    DescriptorIteration(
            @JsonProperty("index") Integer index,
            @JsonProperty("mode") String mode,
            @JsonProperty("merge") Integer merge,
            @JsonProperty("regionName") String regionName,
            @JsonProperty("element") List<DescriptorElement> elements) {
        this.index = index == null ? -1 : index;
        this.merge = merge == null ? 0 : merge;
        this.horizontal = mode == null ? null : "horizontal".equals(mode);
        this.regionName = regionName;
        this.elements = elements;
    }

    public DescriptorIteration(int index, Boolean horizontal, int merge,
                               String regionName) {
        this.index = index;
        this.horizontal = horizontal;
        this.merge = merge;
        this.regionName = regionName;
        this.elements = new LinkedList<>();
    }

    public int getIndex() {
        return index;
    }

    public boolean isHorizontal() {
        return horizontal != null && horizontal;
    }

    public int getMerge() {
        return merge;
    }

    @JsonGetter("element")
    public List<DescriptorElement> getElements() {
        return elements;
    }

    @JsonGetter("index")
    Integer getIndexJSON() {
        return index < 0 ? null : index;
    }

    @JsonGetter("mode")
    String getModeJSON() {
        if (horizontal == null || !horizontal)
            return null;
        return "horizontal";
    }

    @JsonGetter("merge")
    public Integer getMergeJSON() {
        return merge > 0 ? merge : null;
    }

    @JsonGetter("regionName")
    public String getRegionName() {
        return regionName;
    }
}
