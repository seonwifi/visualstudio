/// <reference path="../App.js" />
/// <reference path="../Visualization.js" />

(function () {
    "use strict";

    // 새 페이지가 로드될 때마다 초기화 함수를 실행해야 합니다.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayDataOrRedirect();
        });
    };

    // 바인딩이 있는지 확인하고 시각화를 표시합니다.
    //     또는 [데이터 바인딩] 페이지로 리디렉션합니다.
    function displayDataOrRedirect() {
        Office.context.document.bindings.getByIdAsync(
            app.bindingID,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var binding = result.value;
                    displayDataForBinding(binding);
                    // 그런 다음 변경 이벤트 처리기를 바인딩에 바인딩합니다.
                    binding.addHandlerAsync(
                        Office.EventType.BindingDataChanged,
                        function () {
                            displayDataForBinding(binding);
                        }
                    );
                } else {
                    window.location.href = '../DataBinding/DataBinding.html';
                }
            });
    }

    // 데이터에 대한 바인딩을 쿼리합니다.
    function displayDataForBinding(binding) {
        binding.getDataAsync(
            {
                coercionType: Office.CoercionType.Matrix,
                valueFormat: Office.ValueFormat.Unformatted,
                filterType: Office.FilterType.OnlyVisible
            },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    displayDataHelper(result.value);
                } else {
                    $('#data-display').html(
                        '<div class="notice">' +
                        '    <h2> 데이터를 페치하는 동안 오류가 발생했습니다.</h2>' +
                        '    <a href="../DataBinding/DataBinding.html">' +
                        '        <b>다른 범위에 바인딩하시겠습니까?</b>' +
                        '    </a>' +
                        '</div>');
                }
            }
        );
    }

    // 이미 바인딩을 읽었으므로 데이터를 표시합니다.
    function displayDataHelper(data) {
        var rowCount = data.length;
        var columnCount = (data.length > 0) ? data[0].length : 0;
        if (!visualization.isValidRowAndColumnCount(rowCount, columnCount)) {
            $('#data-display').html(
                '<div class="notice">' +
                '    <h2>데이터가 충분하지 않습니다.</h2>' +
                '    <p>범위에는 다음이 포함되어야 합니다. ' + visualization.rowAndColumnRequirementText + '.</p>' +
                '    <a href="../DataBinding/DataBinding.html">' +
                '        <b>다른 범위를 선택하시겠습니까?</b>' +
                '    </a>' +
                '</div>');
            return;
        }

        var $visualizationContent = visualization.createVisualization(data);

        $('#data-display').empty();
        $('#data-display').append($visualizationContent);
    }
})();