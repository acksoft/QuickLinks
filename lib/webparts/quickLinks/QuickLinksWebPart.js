"use strict";
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var sp_lodash_subset_1 = require("@microsoft/sp-lodash-subset");
var QuickLinks_module_scss_1 = require("./QuickLinks.module.scss");
var strings = require("quickLinksStrings");
require("jquery");
var sp_pnp_js_1 = require("sp-pnp-js");
var QuickLinksWebPart = (function (_super) {
    __extends(QuickLinksWebPart, _super);
    function QuickLinksWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    QuickLinksWebPart.prototype.render = function () {
        if (sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.Local) {
            // this.domElement.innerHTML = this.getLinksHTML();
            this.domElement.innerHTML = "\n      <div class=\"" + QuickLinks_module_scss_1.default.quickLinks + "\">\n        <div class=\"" + QuickLinks_module_scss_1.default.container + "\">\n          <div class=\"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + QuickLinks_module_scss_1.default.row + "\">\n            <div class=\"ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1\">\n              <span class=\"ms-font-xl ms-fontColor-white\">Quick Links (Running Locally)</span>\n              <p class=\"ms-font-l ms-fontColor-white\">" + sp_lodash_subset_1.escape(this.properties.LinkSet) + "</p>\n            </div>\n          </div>\n        </div>\n      </div>";
        }
        else if (sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.SharePoint ||
            sp_core_library_1.Environment.type === sp_core_library_1.EnvironmentType.ClassicSharePoint) {
            this.domElement.innerHTML = "\n      <div class=\"" + QuickLinks_module_scss_1.default.quickLinks + "\">\n        <div class=\"" + QuickLinks_module_scss_1.default.container + "\">\n          <div class=\"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + QuickLinks_module_scss_1.default.row + "\">\n            <div class=\"ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1\">\n              <span class=\"ms-font-xl ms-fontColor-white\">Quick Links (Running in SPO)</span>\n              <p class=\"ms-font-l ms-fontColor-white\">" + sp_lodash_subset_1.escape(this.properties.LinkSet) + " Links</p>\n              <div id=\"qlinks\" />\n            </div>\n          </div>\n        </div>\n      </div>";
            $('#qlinks', this.domElement).html(this.getLinkData());
        }
    };
    Object.defineProperty(QuickLinksWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    QuickLinksWebPart.prototype.getLinkData = function () {
        var linkData = '';
        var linkFilter = "LinkSet eq '" + this.properties.LinkSet + "' and Visible eq 1";
        sp_pnp_js_1.default.sp.web.lists.getByTitle('QuickLinks').items.filter(linkFilter).orderBy("Sequence").get().then(function (r) {
            var row = '';
            // start the unordered list
            linkData = "<ul>";
            // loop through the returned items to create the links
            r.array.forEach(function (item) {
                row += "<li><a href='";
                row += item.URL.Url;
                row += "'";
                // if external, open in new window/tab
                if (item.External) {
                    row += " target='_blank'";
                }
                // if there are comments, turn them into a tooltip
                if (item.Comments.Length > 0) {
                    row += " title='";
                    row += item.Comments;
                    row += "'";
                }
                //set the link text
                row += ">";
                row += item.URL.Description;
                // if extnernal, add the indicator icon
                if (item.External) {
                    row += " <i class='ms-Icon ms-Icon--OpenInNewWindow' aria-hidden='true'></i>";
                }
                // close the link
                row += "</a></li>";
                // add the link to the list of links
                linkData += row;
                row = '';
            });
            // close the link list
            linkData += "</ul>";
        });
        return linkData;
    };
    QuickLinksWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('LinkSet', {
                                    label: strings.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return QuickLinksWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
Object.defineProperty(exports, "__esModule", { value: true });
exports.default = QuickLinksWebPart;

//# sourceMappingURL=QuickLinksWebPart.js.map
