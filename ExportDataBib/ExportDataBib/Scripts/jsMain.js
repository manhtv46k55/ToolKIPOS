function btnSearch() {

    //var list = [];
    //$('#tblData tr').each(function (index, tr) {
    //    $(tr).find('td:eq(1)').each(function (index, td) {
    //        list.push(td.innerHTML);
    //    });
    //});
    window.location.href = "/Home/GetData?iCollectionID=" + $("#CollectionID").val() + "&dtFromDate=" + $("#dtFromDate").val() + "&dtToDate=" + $("#dtToDate").val();
  
    //$.ajax({
    //    url: "/Home/GetData",
    //    data: {
    //        sArrNameField: list.toString()
    //        , iCollectionID: $("#CollectionID").val()
    //        , dtFromDate: $("#dtFromDate").val()
    //        , dtToDate: $("#dtToDate").val()
    //    }
    //});
}