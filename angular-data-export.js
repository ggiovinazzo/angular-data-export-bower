(function () {
    'use strict';

    var app = angular.module('angular.data.export.excel', []);

    function DataTransformExcel() {

        this.mime          = 'application/vnd.ms-excel';
        this.fileExtension = '.xls';
        this.mapping       = null;

        //@TODO Move into a more generic service
        Object.byString = function (o, s) {
            s     = s.replace(/\[(\w+)\]/g, '.$1'); // convert indexes to properties
            s     = s.replace(/^\./, '');           // strip a leading dot
            var a = s.split('.');
            for (var i = 0, n = a.length; i < n; ++i) {
                var k = a[i];
                if (k in o) {
                    o = o[k];
                } else {
                    return;
                }
            }
            return o;
        };

        /**
         * @ngdoc function
         * @name DataTransformExcel.getMimeType
         * @module angular.data.transform.excel
         * @kind function
         * @description Indicates mime type for the transformed data
         * @return {string}
         */
        this.getMimeType = function () {
            return this.mime;
        };

        /**
         * @ngdoc function
         * @name DataTransformExcel.getFileExtension
         * @module angular.data.transform.excel
         * @kind function
         * @description Indicates file extension for the transformed data
         * @return {string}
         */
        this.getFileExtension = function () {
            return this.fileExtension;
        };

        /**
         * @ngdoc function
         * @name DataTransformExcel.setMapping
         * @module angular.data.transform.excel
         * @kind function
         * @description Indicates how to render and map columns
         * @return {string}
         */
        this.setMapping = function (map) {
            this.mapping = map;
        };

        /**
         * @ngdoc function
         * @name DataTransformExcel.processHeader
         * @module angular.data.transform.excel
         * @kind function
         * @description Processes Excel header from json data keys
         * @param data
         * @return {string}
         */
        this.processHeader = function (data) {
            var keys = [];
            // Header
            if (this.mapping == null) {
                if (typeof Object.keys !== 'function') {
                    alert('Cannot transform data, Object.keys function is not available.');
                }
                keys = Object.keys(data[0]);
            } else {
                for (var i = 0; i < this.mapping.length; i++) {
                    keys.push(this.mapping[i].displayName);
                }
            }

            var header = '<tr>';
            for (var j = 0; j < keys.length; j++) {
                header += '<td>' + keys[j] + '</td>';
            }
            header += '</tr>';
            return header;
        };

        /**
         * @ngdoc function
         * @name DataTransformExcel.transform
         * @module angular.data.transform
         * @kind function
         * @description Transforms data into specific format
         * @param {json} data to transform
         * @returns {string}
         */
        this.transform = function (data) {
            var excel = '';
            // Cleanup
            excel += this.processHeader(data);

            //Body
            if (this.mapping == null) {
                for (var i = 0; i < data.length; i++) {
                    excel += '<tr>';
                    for (var cell in data[i]) {
                        excel += '<td>' + data[i][cell] + '</td>';
                    }
                    excel += '</tr>';
                }
            } else {
                for (var k = 0; k < data.length; k++) {
                    excel += '<tr>';
                    for (var j = 0; j < this.mapping.length; j++) {
                        this.mapping[j].field && (excel += '<td>' + Object.byString(data[k], this.mapping[j].field) + '</td>');
                    }
                    excel += '</tr>';
                }
            }

            var excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Angular Export Excel</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>";
            excelFile += excel;
            excelFile += "</table></body></html>";

            return excelFile;
        };

    }

    app.factory('dataExportExcelService',
      function () {
          return new DataTransformExcel();
      }
    );

})();
(function() {
    'use strict';

    var app = angular.module('angular.data.export',
        [
            'base64',
            'angular.download.service'
        ]
    );
    
    function DataTransform($base64,$injector,fileDownloadService){

        this.mapping = null;
        this.module = null;
        
        /**
         * @ngdoc function
         * @name DataTransform.setDataMapping
         * @module angular.data.transform
         * @kind function
         * @description Indicates the field that need to be transformed
         * @param {string} module used to perform data transform
         */
        this.setDataMapping = function(mapping){
            this.mapping = mapping;
        };
        
        /**
         * @ngdoc function
         * @name DataTransform.transform
         * @module angular.data.transform
         * @kind function
         * @description Transforms data into specific format
         * @param {json} data to transform
         * @param {string} module used to perform data transform
         * @returns {string}
         */
        this.transform = function(data,module){
            this.module = $injector.get(module);
            if(this.mapping != null)
                this.module.setMapping(this.mapping);
            return this.module.transform(data);
        };
        
        /**
         * @ngdoc function
         * @name DataTransform.transformAndDownload
         * @module angular.data.transform
         * @kind function
         * @description Transforms data into specific format and launch a download a file action
         * @param {json} data to transform
         * @param {string} module used to perform data transform
         * @param {string} filename for the download
         */
        this.transformAndDownload = function(data,module,filename){
            var tData = this.transform(data,module);
            if(filename==undefined)
                filename = 'exportData' + this.module.getFileExtension();
            fileDownloadService.setMimeType( this.module.getMimeType() );
            fileDownloadService.downloadFile( filename, tData );
        }
        
    }
    
    app.factory('dataExportService', [
        '$base64',
        '$injector',
        'fileDownloadService',
        function($base64,$injector,fileDownloadService){
            return new DataTransform($base64,$injector,fileDownloadService);
        }
    ]);
    
})();