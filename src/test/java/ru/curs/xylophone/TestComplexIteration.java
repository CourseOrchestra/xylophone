package ru.curs.xylophone;

import org.junit.Test;


public class TestComplexIteration extends FullApprovalsTester {

    @Test
    public void testBasicExample() throws XylophoneError {
        approvalTest("complex_iteration/basic_example/descriptor.json",
                "complex_iteration/basic_example/data.xml",
                "complex_iteration/basic_example/template.xlsx",
                OutputType.XLSX, false);
    }


    @Test
    public void testStaticHeader() throws XylophoneError {
        approvalTest("complex_iteration/static_header/descriptor.json",
                "complex_iteration/static_header/data.xml",
                "complex_iteration/static_header/template.xlsx",
                OutputType.XLSX, false);
    }


    @Test
    public void testStaticHeaderAndSidebar() throws XylophoneError {
        approvalTest("complex_iteration/static_header_and_sidebar/descriptor.json",
                "complex_iteration/static_header_and_sidebar/data.xml",
                "complex_iteration/static_header_and_sidebar/template.xlsx",
                OutputType.XLSX, false);
    }

}
