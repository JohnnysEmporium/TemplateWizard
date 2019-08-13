// ==UserScript==
// @name         Template Wizard
// @version      1.1
// @description  Adds a few features to the Service Now console.
// @author       Jan Sobczak
// @match        https://arcelormittalprod.service-now.com/*
// @namespace    https://raw.githubusercontent.com/JohnyHCL/TemplateWizard/master/src/template_copy.js
// @downloadURL  https://raw.githubusercontent.com/JohnyHCL/TemplateWizard/master/src/template_copy.js
// @updateURL	 https://raw.githubusercontent.com/JohnyHCL/TemplateWizard/master/src/template_copy.js
// @grant        GM_setClipboard
// ==/UserScript==

'use strict';

function RUNALL(){

    if(document.readyState === 'complete'){

        clearTimeout(runallTimeout);

        function readText(){

            var incNo = document.getElementById('sys_readonly.incident.number').value;
            console.log(incNo);

            var incStatus
            if(document.getElementById('sys_readonly.incident.incident_state') === null){
                incStatus = document.getElementById('incident.incident_state').options[document.getElementById('incident.incident_state').selectedIndex].text;
            } else {
                incStatus = document.getElementById('sys_readonly.incident.incident_state').options[document.getElementById('sys_readonly.incident.incident_state').selectedIndex].text;

            };
            console.log(incStatus);

            var incPrio = document.getElementById('sys_readonly.incident.priority').value;
            console.log(incPrio);

            var summary = document.getElementById('incident.short_description').value;
            console.log(summary);

            var description = document.getElementById('incident.description').value;
            console.log(description);

            var RG = document.getElementById('sys_display.incident.assignment_group').value;
            console.log(RG);

            var startDate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].lastChild.querySelectorAll('.date-calendar')[0].innerHTML;
            console.log(startDate);

            var latestDate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].firstChild.querySelectorAll('.date-calendar')[0].innerHTML;
            console.log(latestDate);

            var tagName = document.querySelectorAll('.h-card-wrapper')[0].getElementsByTagName('li')[0].childNodes[2].childNodes[0].children[0].tagName;
            var latestUpdate;
            if(tagName == 'SPAN'){
                latestUpdate = document.querySelectorAll('.h-card-wrapper')[0].getElementsByTagName('li')[0].childNodes[2].childNodes[0].childNodes[0].innerHTML;
            } else if(tagName == 'UL'){
                latestUpdate = document.querySelectorAll('.h-card-wrapper')[0].getElementsByTagName('li')[0].childNodes[2].childNodes[0].childNodes[0].childNodes[1].getElementsByTagName('span')[1].childNodes[0].innerHTML
            } else {
                latestUpdate = ''
            };
            console.log(latestUpdate);

            var toPython = [incNo + '/nextEl', incStatus + '/nextEl', incPrio + '/nextEl', summary + '/nextEl', description + '/nextEl', RG + '/nextEl', startDate + '/nextEl', latestDate + '/nextEl', latestUpdate + '/nextEl,'];
            console.log(toPython);
            GM_setClipboard(toPython);
        };

        function latestOnly(){
            var latestDate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].firstChild.querySelectorAll('.date-calendar')[0].innerHTML;
            console.log(latestDate);
            var latestUpdate
            if(document.getElementById('sn_form_inline_stream_entries').childNodes[0].firstChild.querySelectorAll('.sn-widget-textblock-body')[0] === undefined){
                latestUpdate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].firstChild.childNodes[2].childNodes[0].childNodes[0].childNodes[1].childNodes[1].childNodes[0].innerHTML
            } else {
                latestUpdate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].firstChild.querySelectorAll('.sn-widget-textblock-body')[0].innerHTML;
            };
            console.log(latestUpdate);

            var toPython = [latestDate + '/nextEl', latestUpdate + '/nextEl,']
            GM_setClipboard(toPython);
        };

        document.addEventListener('keydown', function(e){
            if (e.keyCode == 53 && !e.shiftKey && !e.ctrlKey && e.altKey && !e.metaKey) {
                readText();
            }
        }, false);

        document.addEventListener('keydown', function(e){
            if (e.keyCode == 54 && !e.shiftKey && !e.ctrlKey && e.altKey && !e.metaKey) {
                latestOnly();
            }
        }, false);


    } else {
        var runallTimeout = setTimeout(RUNALL, 300);
        runallTimeout;
    };
};
RUNALL();
