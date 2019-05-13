<%@ page language="java" pageEncoding="UTF-8"%>
<%@ page contentType="text/html;charset=UTF-8"%>

<a href="excel/export">导出</a> <br/>
<form action="excel/import" enctype="multipart/form-data" method="post">
    <input type="file" name="file"/>
    <input type="submit" value="导入Excel">
</form>