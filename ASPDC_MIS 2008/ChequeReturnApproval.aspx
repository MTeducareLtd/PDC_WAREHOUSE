<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="ChequeReturnApproval.aspx.vb" 
Inherits="ChequeReturnApproval" title="Cheque Return Approval" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div id="main-content" class="clearfix">
        <div id="breadcrumbs">
            <ul class="breadcrumb">
                <li><i class="icon-home"></i><a href="UserDashboard.aspx">Home</a><span class="divider"><i
                    class="icon-angle-right"></i></span></li>
                <li class="active">Cheque Return Approval</li>
            </ul>
            <!--.breadcrumb-->
            
        </div>
        <!--end of Breadcrumb-->
        <div id="page-content" class="clearfix">
            <div class="row-fluid">
                <h3 class="header smaller lighter blue">
                    Cheque Return Approval
                    <div class="btn-group" style="position: absolute; left: 87%">
                        <asp:Button class="btn btn-small btn-primary" ID="btnSearch" runat="server" Text="Search" />
                        
                    </div>
                </h3>
                <div class="row-fluid" id="DivSearch" visible="False" runat="server">
                    <div class="span12">
                        <div class="widget-box">
                            <div class="widget-header widget-header-small header-color-dark">
                                <h6>
                                    Search Options</h6>
                                <div class="widget-toolbar">
                                            <a href="#" data-action="close" class="ace-tooltip" placeholder="Tooltip on hover"
                                            title="" data-placement="middle" data-original-title="Close"><i class="icon-remove">
                                            </i></a>

                                </div>
                            </div>
                            <div class="widget-body">
                                <div class="widget-body-inner">
                                    <div class="widget-main padding-3">
                                        
                                            <div class="slim-scroll">
                                                <table id="table2" runat="server">
                                                    <tr id="Tr2" runat="server">
                                                        <td>
                                                            <div class="row-fluid">
                                                                <label for="ddlDivision">Division Name<small class="text-warning">*</small></label>
						                                        
                                                                    <asp:ListBox ID="ddlDivision" class="chzn-select" 
                                                                        data-placeholder="Search Division..." runat="server" 
                                                                    AutoPostBack="True">
                                                                    </asp:ListBox>
						                                       

                                                            </div>
                                                            
                                                        </td>
                                                        <td>
                                                            <div class="row-fluid">
                                                                <label for="ddlCentre">Request Type</label>
						                                        <%--<div class="input-append">--%>
							                                        <asp:ListBox ID="ddlRequestType"  runat="server" class="chzn-select" 
                                                                     data-placeholder="Request Type..." SelectionMode="Single" >
                                                                     <asp:ListItem>Pending</asp:ListItem> 
                                                                     <asp:ListItem>Approved</asp:ListItem> 
                                                                     <asp:ListItem>Rejected</asp:ListItem> 
                                                                     <asp:ListItem>All</asp:ListItem> 
                                                                    </asp:ListBox>
                                                                    
						                                        <%--</div>--%>

                                                            </div>
                                                            
                                                        </td>
                                                        <td>
                                                            <div class="row-fluid">
                                                                <label for="ddlCentre">Period</label>
						                                        <div class="input-prepend">
			                                                            <span class="add-on"><i class="icon-calendar"></i></span>
                                                                        <input runat="server"  class="span11 ace-tooltip" name="date-range-picker" id="id_date_range_picker_1" placeholder="Date Search" data-placement="bottom" data-original-title="Search by Date Range"/>
			                                                            
			                                                        </div>
                                                                    
                                                            </div>
                                                            
                                                        </td>
                                                        <td>
                                                            <div class="row-fluid">
                                                                 &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button class="btn btn-small btn-primary" ID="btnSearchRecord" runat="server" Text="Search" /> 
                                                            </div>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td colspan="4">
                                                        </td>
                                                    </tr>
                                                    
                                                    
                                                </table>
                                            </div>

                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                 <div class="alert alert-block alert-success" id="MsgSuccess" visible="false" runat ="server">
			            <button type="button" class="close" data-dismiss="alert"><i class="icon-remove"></i></button>
			                <p>
			                <strong><i class="icon-ok"></i> Well done!</strong>
                              <asp:Label ID="lblSuccess" runat="server" Text="Label"></asp:Label>
			                </p>
			
		            </div>
                <div class="table-header" id="Searchresult" visible="false" runat="server">
                    Result</div>
                <div id="table_report_wrapper" class="dataTables_wrapper" role="grid" visible="false"
                    runat="server">
                    <div class="row-fluid">
                        <div class="span9">
                            <div id="table_report_length" class="dataTables_length">
                                <label>
                                    Total No. of rows:&nbsp;
                                    <asp:Label ID="lblRowCnt" runat="server" Text='0'></asp:Label>
                                </label>
                               
                            </div>
                        </div>
                        <div class="span3" align="right">
                            <asp:Button class="btn btn-small btn-success" ID="btnExport" runat="server" Text="Export to Excel" />
                        </div> 
                    </div>
                    <br />

                    <div class="pull-right center" id="spinner_preview" runat ="server" ></div>

                    <asp:DataList ID="dlReport" class="table table-striped table-bordered table-hover" runat="server" Width="100%">
                        <HeaderTemplate>
                                
                                        Centre
                                    </th>
                                    <th style="width: 10%" align="left">
                                        Stream
                                    </th>
                                    <th style="width: 15%" align="left">
                                        Student Name
                                    </th>
                                    <th style="width: 10%" align="left">
                                        Cheque No
                                    </th>
                                    <th style="width: 10%" align="left">
                                        Cheque Date
                                    </th>
                                    <th style="width: 10%" align="left">
                                        Amount
                                    </th>
                                    <th style="width: 10%" align="left">
                                        Request Date
                                    </th>
                                    <th style="width: 15%" align="left">
                                        Request Reason
                                    </th>
                                    <th style="width: 10%" align="left">
                                        Action
                        </HeaderTemplate>
                        <ItemTemplate>
                                
                                        <asp:Label ID="lblName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"centername" )%>'></asp:Label>
                                    </td>
                                    <td style="width: 10%" align="left">
                                        <asp:Label ID="Label1" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"streamname" )%>'></asp:Label>
                                    </td>
                                    <td style="width: 15%" align="left">
                                        <asp:Label ID="lblurl" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"studentname")%>'></asp:Label>
                                    </td>
                                    <td style="width: 10%" align="left">
                                        <asp:Label ID="Label3" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CenterChequeNo")%>'></asp:Label>
                                    </td>
                                    <td style="width: 10%" align="left">
                                        <asp:Label ID="Label4" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CentreChequeDate")%>'></asp:Label>
                                    </td>
                                    <td style="width: 10%" align="left">
                                        <asp:Label ID="Label5" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CentreChequeAmt")%>'></asp:Label>
                                    </td>
                                    <td style="width: 10%" align="left">
                                        <asp:Label ID="Label6" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"RequestDate")%>'></asp:Label>
                                    </td>
                                    <td style="width: 15%" align="left">
                                        <asp:Label ID="Label7" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ReturnReasonNote")%>'></asp:Label>
                                    </td>
                                    <td style="width: 10%" align="left">
                                        <asp:Button class="btn btn-small btn-success" ID="btnApprove" runat="server" Text="Approve" CommandName="Approve" Enabled='<%#DataBinder.Eval(Container.DataItem,"ReturnStatus")%>' CommandArgument='<%#DataBinder.Eval(Container.DataItem,"ReturnRequestCode") %>' />
                                        <asp:Button class="btn btn-small btn-danger" ID="btnReject" runat="server" Text="Reject"  CommandName="Reject" Enabled='<%#DataBinder.Eval(Container.DataItem,"ReturnStatus")%>' CommandArgument='<%#DataBinder.Eval(Container.DataItem,"ReturnRequestCode") %>' />
                        </ItemTemplate>
                    </asp:DataList>

                    
                </div>

            </div>
            <!--/row-->
        </div>
        <!--/#page-content-->
    </div>
    <!-- #main-content -->

    		<script type="text/javascript">
		
		$(function() {
			$('#id-disable-check').on('click', function() {
				var inp = $('#form-input-readonly').get(0);
				if(inp.hasAttribute('disabled')) {
					inp.setAttribute('readonly' , 'true');
					inp.removeAttribute('disabled');
					inp.value="This text field is readonly!";
				}
				else {
					inp.setAttribute('disabled' , 'disabled');
					inp.removeAttribute('readonly');
					inp.value="This text field is disabled!";
				}
			});
		
		
			$(".chzn-select").chosen(); 
			$(".chzn-select-deselect").chosen({allow_single_deselect:true}); 
			
			$('.ace-tooltip').tooltip();
			$('.ace-popover').popover();
			
			$('textarea[class*=autosize]').autosize({append: "\n"});
			$('textarea[class*=limited]').each(function() {
				var limit = parseInt($(this).attr('data-maxlength')) || 100;
				$(this).inputlimiter({
					"limit": limit,
					remText: '%n character%s remaining...',
					limitText: 'max allowed : %n.'
				});
			});
			
			$.mask.definitions['~']='[+-]';
			$('.input-mask-date').mask('99/99/9999');
			$('.input-mask-phone').mask('(999) 999-9999');
			$('.input-mask-eyescript').mask('~9.99 ~9.99 999');
			$(".input-mask-product").mask("a*-999-a999",{placeholder:" ",completed:function(){alert("You typed the following: "+this.val());}});
			
			
			
			$( "#input-size-slider" ).css('width','200px').slider({
				value:1,
				range: "min",
				min: 1,
				max: 6,
				step: 1,
				slide: function( event, ui ) {
					var sizing = ['', 'input-mini', 'input-small', 'input-medium', 'input-large', 'input-xlarge', 'input-xxlarge'];
					var val = parseInt(ui.value);
					$('#form-field-4').attr('class', sizing[val]).val('.'+sizing[val]);
				}
			});

			$( "#input-span-slider" ).slider({
				value:1,
				range: "min",
				min: 1,
				max: 11,
				step: 1,
				slide: function( event, ui ) {
					var val = parseInt(ui.value);
					$('#form-field-5').attr('class', 'span'+val).val('.span'+val).next().attr('class', 'span'+(12-val)).val('.span'+(12-val));
				}
			});
			
			
			var $tooltip = $("<div class='tooltip right in' style='display:none;'><div class='tooltip-arrow'></div><div class='tooltip-inner'></div></div>").appendTo('body');
			$( "#slider-range" ).css('height','200px').slider({
				orientation: "vertical",
				range: true,
				min: 0,
				max: 100,
				values: [ 17, 67 ],
				slide: function( event, ui ) {
					var val = ui.values[$(ui.handle).index()-1]+"";
					
					var pos = $(ui.handle).offset();
					$tooltip.show().children().eq(1).text(val);		
					$tooltip.css({top:pos.top - 10 , left:pos.left + 18});
					
					//$(this).find('a').eq(which).attr('data-original-title' , val).tooltip('show');
				}
			});
			$('#slider-range a').tooltip({placement:'right', trigger:'manual', animation:false}).blur(function(){
				$tooltip.hide();
				//$(this).tooltip('hide');
			});
			//$('#slider-range a').tooltip({placement:'right', animation:false});
			
			$( "#slider-range-max" ).slider({
				range: "max",
				min: 1,
				max: 10,
				value: 2,
				//slide: function( event, ui ) {
				//	$( "#amount" ).val( ui.value );
				//}
			});
			//$( "#amount" ).val( $( "#slider-range-max" ).slider( "value" ) );
			
			$( "#eq > span" ).css({width:'90%', float:'left', margin:'15px'}).each(function() {
				// read initial values from markup and remove that
				var value = parseInt( $( this ).text(), 10 );
				$( this ).empty().slider({
					value: value,
					range: "min",
					animate: true
					
				});
			});

			
			$('#id-input-file-1 , #id-input-file-2').ace_file_input({
				no_file:'No File ...',
				btn_choose:'Choose',
				btn_change:'Change',
				droppable:false,
				onchange:null,
				thumbnail:false //| true | large
				//whitelist:'gif|png|jpg|jpeg'
				//blacklist:'exe|php'
				//onchange:''
				//
			});
			
			$('#id-input-file-3').ace_file_input({
				style:'well',
				btn_choose:'Drop files here or click to choose',
				btn_change:null,
				no_icon:'icon-cloud-upload',
				droppable:true,
				onchange:null,
				thumbnail:'small',
				before_change:function(files, dropped) {
					/**
					if(files instanceof Array || (!!window.FileList && files instanceof FileList)) {
						//check each file and see if it is valid, if not return false or make a new array, add the valid files to it and return the array
						//note: if files have not been dropped, this does not change the internal value of the file input element, as it is set by the browser, and further file uploading and handling should be done via ajax, etc, otherwise all selected files will be sent to server
						//example:
						var result = []
						for(var i = 0; i < files.length; i++) {
							var file = files[i];
							if((/^image\//i).test(file.type) && file.size < 102400)
								result.push(file);
						}
						return result;
					}
					*/
					return true;
				}
				/*,
				before_remove : function() {
					return true;
				}*/

			}).on('change', function(){
				//console.log($(this).data('ace_input_files'));
				//console.log($(this).data('ace_input_method'));
			});

			
			$('#spinner1').ace_spinner({value:0,min:0,max:200,step:10, btn_up_class:'btn-info' , btn_down_class:'btn-info'})
			.on('change', function(){
				//alert(this.value)
			});
			$('#spinner2').ace_spinner({value:0,min:0,max:10000,step:100, icon_up:'icon-caret-up', icon_down:'icon-caret-down'});
			$('#spinner3').ace_spinner({value:0,min:-100,max:100,step:10, icon_up:'icon-plus', icon_down:'icon-minus', btn_up_class:'btn-success' , btn_down_class:'btn-danger'});

			
			/**
			var stubDataSource = {
			data: function (options, callback) {

				setTimeout(function () {
					callback({
						data: [
							{ name: 'Test Folder 1', type: 'folder', additionalParameters: { id: 'F1' } },
							{ name: 'Test Folder 1', type: 'folder', additionalParameters: { id: 'F2' } },
							{ name: 'Test Item 1', type: 'item', additionalParameters: { id: 'I1' } },
							{ name: 'Test Item 2', type: 'item', additionalParameters: { id: 'I2' } }
						]
					});
				}, 0);

			}
			};
			$('#MyTree').tree({ dataSource: stubDataSource , multiSelect:true })
			*/
			
			$('.date-picker').datepicker();
			$('#timepicker1').timepicker({
				minuteStep: 1,
				showSeconds: true,
				showMeridian: false
			});
			
			$(' #<%= id_date_range_picker_1.ClientID %>').daterangepicker();
			
			$('#colorpicker1').colorpicker();
			$('#simple-colorpicker-1').ace_colorpicker();
		
			
		$(".knob").knob();
	

		});




		</script>
</asp:Content>

