// ==UserScript==
// @name         Template_Copy
// @version      1.0
// @description  Adds a few features to the Service Now console.
// @author       Jan Sobczak
// @match        https://arcelormittalprod.service-now.com/*
// @grant        GM_setClipboard
// ==/UserScript==

'use strict';

function RUNALL(){

    if(document.readyState === 'complete'){

        clearTimeout(runallTimeout);

        function readText(){
            var incNo = document.getElementById('sys_readonly.incident.number').value;
            var incStatus = document.getElementById('incident.incident_state').options[document.getElementById('incident.incident_state').selectedIndex].text;
            var incPrio = document.getElementById('sys_readonly.incident.priority').value
            var summary = document.getElementById('sys_readonly.incident.short_description').value;
            var description = document.getElementById('incident.description').value;
            var RG = document.getElementById('sys_display.incident.assignment_group').value;
            var startDate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].lastChild.querySelectorAll('.date-calendar')[0].innerHTML;
            var latestDate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].firstChild.querySelectorAll('.date-calendar')[0].innerHTML;
            var latestUpdate = document.getElementById('sn_form_inline_stream_entries').childNodes[0].firstChild.querySelectorAll('.sn-widget-textblock-body')[0].innerHTML;

            var toPython = [incNo + '/nextEl', incStatus + '/nextEl', incPrio + '/nextEl', summary + '/nextEl', description + '/nextEl', RG + '/nextEl', startDate + '/nextEl', latestDate + '/nextEl', latestUpdate + '/nextEl,'];
            console.log(toPython)
            GM_setClipboard(toPython)

//             var dict = {
//                 incNo: incNo,
//                 incStatus: incStatus,
//                 incPrio: incPrio,
//                 summary: summary,
//                 description: description,
//                 RG: RG,
//                 startDate: startDate,
//                 latestUpdate: latestUpdate,
//             };

//             console.log(dict)
            GM_setClipboard(toPython)

        };

        document.addEventListener('keydown', function(e){
            if (e.keyCode == 53 && !e.shiftKey && !e.ctrlKey && e.altKey && !e.metaKey) {
                readText();
            }
        }, false);


    } else {
        var runallTimeout = setTimeout(RUNALL, 300);
        runallTimeout;
    };
};
RUNALL();
