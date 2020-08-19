(function (window, location) {
    "use strict";

    switch (String(location.hostname)) {
        case "products.groupdocs.app":
            window.GROUPDOCS_ASSEMBLY_API_BASE = "/GroupDocsAPI/api/GroupDocsAssembly/";
            break;
        case "products-qa.groupdocs.app":
            window.GROUPDOCS_ASSEMBLY_API_BASE = "/GroupDocsAPI/api/GroupDocsAssembly/";
            break;
        case "localhost":
            window.GROUPDOCS_ASSEMBLY_API_BASE = "http://localhost:2122/api/GroupDocsAssembly/";
            break;
    }
})(window, location);

