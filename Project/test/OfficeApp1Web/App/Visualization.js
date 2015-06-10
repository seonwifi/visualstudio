var visualization = (function () {
    "use strict";

    var visualization = {};

    // 샘플 데이터:
    visualization.sampleHeaders = [['이름', '등급']];
    visualization.sampleRows = [
        ['Ben', 79],
        ['Amy', 95],
        ['Jacob', 86],
        ['Ernie', 93]];

    // 데이터 범위 유효성 검사:
    visualization.rowAndColumnRequirementText = '2개의 열과 최소 2개의 행';
    visualization.isValidRowAndColumnCount = function (rowCount, columnCount) {
        return (rowCount > 1 && columnCount === 2);
    };

    // 전달된 데이터를 기반으로 시각화를 만듭니다.
    visualization.createVisualization = function (data) {
        var maxBarWidthInPixels = 200;

        var $table = $('<table class="visualization" />');
        var $headerRow = $('<tr />').appendTo($table);
        $('<th />').text(data[0][0]).appendTo($headerRow);
        $('<th />').text(data[0][1]).appendTo($headerRow);

        for (var i = 1; i < data.length; i++) {
            var $row = $('<tr />').appendTo($table);
            var $column1 = $('<td />').appendTo($row);
            var $column2 = $('<td />').appendTo($row);

            $column1.text(data[i][0]);
            var value = data[i][1];
            var width = (maxBarWidthInPixels * value / 100.0);
            var $visualizationBar = $('<div />').appendTo($column2);
            $visualizationBar.addClass('bar')
                .width(width)
                .text(value);
        }

        return $table;
    };

    return visualization;
})();