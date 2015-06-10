/// <reference path="../App.js" />
/// <reference path="../Visualization.js" />

(function () {
    "use strict";

    // 새 페이지가 로드될 때마다 초기화 함수를 실행해야 합니다.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#insert-sample-data').click(insertSampleData);
            $('#bind-to-existing-data').click(bindToExistingData);
        });
    };

    function insertSampleData() {
        var sampleData = new Office.TableData(
            visualization.sampleRows,
            visualization.sampleHeaders);
        Office.context.document.setSelectedDataAsync(sampleData,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    app.showNotification('샘플 데이터를 삽입할 수 없습니다.',
                        '다른 선택 범위를 선택하십시오.');
                } else {
                    Office.context.document.bindings.addFromSelectionAsync(
                        Office.BindingType.Table, { id: app.bindingID },
                        function (asyncResult) {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                app.showNotification('데이터를 바인딩하는 동안 오류가 발생했습니다.');
                            } else {
                                window.location.href = '../Home/Home.html';
                            }
                        }
                    );
                }
            }
        );
    }

    function bindToExistingData() {
        Office.context.document.bindings.addFromSelectionAsync(
            Office.BindingType.Matrix,
            { id: app.bindingID },
            function (result) {
                var isValid = (result.status == Office.AsyncResultStatus.Succeeded) &&
                    visualization.isValidRowAndColumnCount(
                        result.value.rowCount, result.value.columnCount);
                if (isValid) {
                    window.location.href = '../Home/Home.html';
                } else {
                    app.showNotification('잘못된 데이터를 선택했습니다.',
                        '다른 것을 선택하고 다음을 선택했는지 확인하십시오. ' +
                        '다음이 포함된 표 또는 범위 ' + visualization.rowAndColumnRequirementText);
                }
            }
        );
    }
})();