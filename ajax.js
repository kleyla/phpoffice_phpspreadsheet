function ajax(){
    console.log('EXCEL');

    $.ajax({
        success: function(){
            window.open('{{route("excelPresupuesto", ["id" => "1"])}}','_blank' );
        },

    });
}
