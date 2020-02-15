(function ($) {
    function Tree() {
        var $this = this;
        function treeNodeClick() {
            $(document).on('click', '.tree li a input[type="checkbox"]', function () {
                $(this).closest('li').find('ul input[type="checkbox"]').prop('checked', $(this).is(':checked'));
            }).on('click', '.node-item', function () {
                var parentNode = $(this).parents('.tree ul');
                if ($(this).is(':checked')) {
                    parentNode.find('li a .parent').prop('checked', true);
                } else {
                    var elements = parentNode.find('ul input[type="checkbox"]:checked');
                    if (elements.length === 0) {
                        parentNode.find('li a .parent').prop('checked', false);
                    }
                }
            });
        };
        $this.init = function () {
            treeNodeClick();
        }
    }
    $(function () {
        var self = new Tree();
        self.init();
    })
}(jQuery));
$(function () {
    $.ajax({
        type: 'get',
        dataType: 'json',
        cache: false,
        url: '/Home/GetTreeViewSample',
        data: {'target' : $('#targetTreeView').val()},
        success:
            function (json) {
                $('#TreeViewSample')
                    .jstree
                    ({
                        'core':
                        {
                            'data': json
                        },
                        "plugins": ["search", "checkbox", "types"],
                        "types":
                        {
                            "default": { "icon": "glyphicon glyphicon-flash" },
                            "demo": { 'valid_children': ['none'], "icon": "glyphicon glyphicon-question-sign" }
                        }
                    })
                    .on('changed.jstree',
                    function (e, data) {
                        var i, j, r = [];
                        for (i = 0, j = data.selected.length; i < j; i++) {
                            var TargetNode = data.instance.get_node(data.selected[i]);
                            var firsChildren = data.instance.get_node(TargetNode.children[0]);
                            if (TargetNode.children.length === 0) {
                                var wId = data.instance.get_node(data.selected[i]).data;
                                if (jQuery.inArray(wId, r) === -1) {
                                    r.push(wId);
                                }
                            }
                        }
                        var idList = r.join(';');
                        $('#CID_Sample').val(idList);
                    }
                    );                
            }
    });
});