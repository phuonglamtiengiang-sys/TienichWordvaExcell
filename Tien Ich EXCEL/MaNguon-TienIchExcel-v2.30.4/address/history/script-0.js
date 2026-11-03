// JScript File
function TextCombo(tb,e,cmb,field) {            
            var a = tb.GetText();           
            var c = GetItemByField(cmb,a,field)            
            var f = cmb.GetText();
            var r = null;
            if (c != null){
                cmb.SetSelectedIndex(c);
                r = cmb.GetValue();
                return r;
            }
            else {  
                    if (f =='')     {
                          tb.SetText('')
                    }      
                    else {
                        var text = cmb.GetSelectedItem(cmb.GetSelectedIndex).GetColumnText(field);
                        tb.SetText(text);
                    }
                    r = null;
                    alert('Mã đơn vị không có trong danh mục.');
                    return r;                    
            }
        }
        function TextCB(tb,cmb,field) {            
            var a = tb.GetText();           
            var c = GetItemByField(cmb,a,field)            
            var f = cmb.GetText();
            var r = null;
            if (c != null){
                cmb.SetSelectedIndex(c);
                r = cmb.GetValue();
                return r;
            }
            else {  
                    if (f =='')     {
                          tb.SetText('')
                    }      
                    else {
                        var text = cmb.GetSelectedItem(cmb.GetSelectedIndex).GetColumnText(field);
                        tb.SetText(text);
                    }
                    r = null;
                    alert('Mã đơn vị không có trong danh mục.');
                    return r;                    
            }
        }
        
function ComboText(cmb, e, tb,field) {
            var a = cmb.GetSelectedItem(cmb.GetSelectedIndex()).GetColumnText(field);
            tb.SetText(a);            
        }        

function txtChange(tb,e,cmb,field) {
            
            var a = tb.GetText();           
            var c = GetItemByField(cmb,a,field)            
            var f = cmb.GetText();
            if (c != null){
                cmb.SetSelectedIndex(c);
                e.processOnServer = true;
            }
            else {  
                    if (f =='')     {
                          tb.SetText('')
                    }      
                    else {
                        var text = cmb.GetSelectedItem(cmb.GetSelectedIndex).GetColumnText(field);
                        tb.SetText(text);
                    }                    
                    alert('Mã đơn vị không có trong danh mục.');
                    e.processOnServer = false                
            }
        }
        
        function GetItemByField(combo, value,field) {
          for(var i = 0; i < combo.GetItemCount(); i++)
            if(combo.GetItem(i).GetColumnText(field) == value) 
              return i;
          return null;
        }
function cmbChange(cmb, e, tb,field) {
            var a = cmb.GetSelectedItem(cmb.GetSelectedIndex()).GetColumnText(field);
            tb.SetText(a);
            e.processOnServer = true
        }
        function GetItemByValue(combo, value) {
          for(var i = 0; i < combo.GetItemCount(); i++)
            if(combo.GetItem(i).value == value) 
              return combo.GetItem(i);
          return null;
}
function txtChangeClient(tb,e,cmb,field){
    var a = tb.GetText();   
    var c = GetItemByField(cmb,a,field)            
            var f = cmb.GetText();
            if (c != null){
                cmb.SetSelectedIndex(c);
               // e.processOnServer = true;
            }
            else {  
                    if (f =='')     {
                          tb.SetText('')
                    }      
                    else {
                        var text = cmb.GetSelectedItem(cmb.GetSelectedIndex).GetColumnText(field);
                        tb.SetText(text);
                    }                    
                    alert('Mã đơn vị không có trong danh mục.');
                    e.processOnServer = false                
            }
}
function cmbChangeClient(cmb, e, tb,field){
    var a = cmb.GetSelectedItem(cmb.GetSelectedIndex()).GetColumnText(field);
            tb.SetText(a);
}




