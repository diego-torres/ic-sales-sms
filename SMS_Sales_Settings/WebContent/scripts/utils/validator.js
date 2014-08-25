	jQuery(function() {
		jQuery('body').keydown(function() {
		   hideScrim();    
	    });
	});

    function hideScrim() {
        jQuery('#scrim').hide();
        jQuery('#messageDisplayer').hide();
    }
    function validateEmail(email) {
        var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
        return re.test(email);
    }
    function validate(e) {
        var targetMessage = jQuery('#messageDisplayer').text('');
        var strErrorMessage = '';
        var fieldNeedsCorrection = null;
        if (!e) e = window.event;
        var target = e.target || e.srcElement;
        jQuery('[validate]').filter('[validationid = ' + target.attributes['validationid'].value + ']').each(function () {
            if (jQuery(this).css('display') != "none" && jQuery(this).is(":visible")) {
                jQuery(this).val(jQuery.trim(jQuery(this).val()));
                switch (jQuery(this).attr('validate')) {
                    case 'number':
                        if (jQuery(this).val() == '')
                            jQuery(this).val(0);
                        if (isNaN(jQuery(this).val()) || Number(jQuery(this).val()) < 0) {
                            strErrorMessage = 'Number did not pass validation.';
                            fieldNeedsCorrection = jQuery(this);
                            return false;
                        }
                        break;
                    case 'required':
                        if (jQuery(this).val() == '') {
                            strErrorMessage = 'Field is required.';
                            fieldNeedsCorrection = jQuery(this);
                            return false;
                        }
                        break;
                    case 'email':
                        if (!validateEmail(jQuery(this).val())) {
                            strErrorMessage = 'Email did not pass validation.';
                            fieldNeedsCorrection = jQuery(this);
                            return false;
                        }
                        break;
                    case 'numbers':
                        if (jQuery(this).val() == '')
                            jQuery(this).val(0);
                        var arrNumbers = jQuery(this).val();
                        if (arrNumbers != '') {
                            arrNumbers = arrNumbers.split(',');
                            for (var i = 0; i < arrNumbers.length; i++) {
                                arrNumbers[i] = jQuery.trim(arrNumbers[i]);
                                if (isNaN(arrNumbers[i]) || Number(arrNumbers[i]) < 0) {
                                    strErrorMessage = 'Number did not pass validation.';
                                    fieldNeedsCorrection = jQuery(this);
                                    return false;
                                }
                            }
                        }
                        else {
                            strErrorMessage = 'Number did not pass validation.';
                            fieldNeedsCorrection = jQuery(this);
                            return false;
                        }
                        break;
                    case 'date':
                        var dateToValidate = new Date(jQuery(this).val());
                        dateToValidate = dateToValidate.toJSON();
                        if (dateToValidate == null) {
                            strErrorMessage = 'Date did not pass validation.';
                            fieldNeedsCorrection = jQuery(this);
                            return false;
                        }
                        break;
                    case 'dateOrEmpty':
                        if (jQuery(this).val() != '') {
                            var dateToValidate = new Date(jQuery(this).val());
                            dateToValidate = dateToValidate.toJSON();
                            if (dateToValidate == null) {
                                strErrorMessage = 'Date did not pass validation.';
                                fieldNeedsCorrection = jQuery(this);
                                return false;
                            } 
                        }
                        break;
                    default:
                }
            }
        });

        if (strErrorMessage != '') {
            jQuery('#scrim').show();
            targetMessage.text(strErrorMessage).show();

            var pos = fieldNeedsCorrection.offset();
            targetMessage.css('left', (pos.left - 9) + 'px');
            targetMessage.css('top', (pos.top + fieldNeedsCorrection.outerHeight() + 2) + 'px');
            jQuery('html, body').animate({
                scrollTop: fieldNeedsCorrection.offset().top - 300
            }, 100);
            fieldNeedsCorrection.focus().select();
            return false;
        }
        return true;
    }