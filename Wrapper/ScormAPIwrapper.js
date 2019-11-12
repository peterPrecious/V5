  var vDebug_All

  // Scorm 1.2 API
  function API() {

    if (vDebug_All) alert ("entering main API...");

  	//data members
  	this.APIInitialized = false;
  	this.LastErrorNum = "301";
  
  	this.elements = new Array();
  	this.elements[0]  = "cmi.core._children";
  	this.elements[1]  = "cmi.core.score._children";
  	this.elements[2]  = "cmi.core.student_id";
  	this.elements[3]  = "cmi.core.student_name";
  	this.elements[4]  = "cmi.core.lesson_location";
  	this.elements[5]  = "cmi.core.credit";
  	this.elements[6]  = "cmi.core.lesson_status";
  	this.elements[7]  = "cmi.core.entry";
  	this.elements[8]  = "cmi.core.score.raw";
  	this.elements[9]  = "cmi.core.score.max";		//not mandatory
  	this.elements[10] = "cmi.core.total_time";
  	this.elements[11] = "cmi.core.exit";
  	this.elements[12] = "cmi.core.session_time";
  	this.elements[13] = "cmi.suspend_data";
  	this.elements[14] = "cmi.launch_data";
  	this.elements[15] = "cmi.student_data.mastery_score";
  	this.elements[16] = "cmi.student_data.max_time_allowed";
  	this.elements[17] = "cmi.student_data.time_limit_action";
    this.elements[18] = "cmi.interactions";


  	// parallel elements array for looping thru values array in the application
  	this.dup_elements = new Array();
  	this.dup_elements[0]  = "cmi_core_children";
  	this.dup_elements[1]  = "cmi_core_score_children";
  	this.dup_elements[2]  = "cmi_core_student_id";
  	this.dup_elements[3]  = "cmi_core_student_name";
  	this.dup_elements[4]  = "cmi_core_lesson_location";
  	this.dup_elements[5]  = "cmi_core_credit";
  	this.dup_elements[6]  = "cmi_core_lesson_status";
  	this.dup_elements[7]  = "cmi_core_entry";
  	this.dup_elements[8]  = "cmi_core_score_raw";
  	this.dup_elements[9]  = "cmi_core_score_max";		//not mandatory
  	this.dup_elements[10] = "cmi_core_total_time";
  	this.dup_elements[11] = "cmi_core_exit";
  	this.dup_elements[12] = "cmi_core_session_time";
  	this.dup_elements[13] = "cmi_suspend_data";
  	this.dup_elements[14] = "cmi_launch_data";
  	this.dup_elements[15] = "cmi_student_data_mastery_score";
  	this.dup_elements[16] = "cmi_student_data_max_time_allowed";
  	this.dup_elements[17] = "cmi_student_data_time_limit_action";
    this.dup_elements[18] = "cmi.interactions";
   
 
  	this.values = new Array();
  	this.values[0]  = "student_id,student_name,lesson_location,credit,lesson_status,entry,score,total_time,exit,session_time";
  	this.values[1]  = "raw,max";
  	this.values[2]  = ""; // "<%=Session("MembId")%>";
  	this.values[3]  = ""; // "<%=Session("MembFirstName") & " " & Session("MembLastName")%>";
  	this.values[4]  = "";
  	this.values[5]  = "no-credit";
  	this.values[6]  = "not attempted";
  	this.values[7]  = "";
  	this.values[8]  = "";
  	this.values[9]  = "";
  	this.values[10] = "00:00:00";
  	this.values[11] = "";
  	this.values[12] = "00:00:00";
  	this.values[13] = "";
  	this.values[14] = "";
  	this.values[15] = ""; //cmi.student_data.mastery_score
  	this.values[16] = ""; //cmi.student_data.max_time_allowed
  	this.values[17] = ""; //cmi_student_data_time_limit_action
    this.values[18] = ""; 


  	this.errCodes = new Array();
  	this.errCodes["0"]   = "No Error";
  	this.errCodes["101"] = "General Exception";
  	this.errCodes["201"] = "Invalid Argument Error";
  	this.errCodes["202"] = "Element cannot have children";
  	this.errCodes["203"] = "Element not an array.  Cannot have count";
  	this.errCodes["301"] = "Not initialized";
  	this.errCodes["401"] = "Not implemented error";
  	this.errCodes["402"] = "Invalid set value, element is a keyword";
  	this.errCodes["403"] = "Element is read only";
  	this.errCodes["404"] = "Element is write only";
  	this.errCodes["405"] = "Incorrect Data Type";
  
  	this.errDiagn = new Array();
  	this.errDiagn["0"]   = "No Error";
  	this.errDiagn["101"] = "Possible Server error.  Contact System Administrator";
  	this.errDiagn["201"] = "The course made an incorrect function call.  Contact course vendor or system administrator";
  	this.errDiagn["202"] = "The course made an incorrect data request. Contact course vendor or system administrator";
  	this.errDiagn["203"] = "The course made an incorrect data request. Contact course vendor or system administrator";
  	this.errDiagn["301"] = "The system has not been initialized correctly.  Please contact your system administrator";
  	this.errDiagn["401"] = "The course made a request for data not supported by Answers.";
  	this.errDiagn["402"] = "The course made a bad data saving request.  Contact course vendor or system adminsitrator";
  	this.errDiagn["403"] = "The course tried to write to a read only value.  Contact course vendor";
  	this.errDiagn["404"] = "The course tried to read a value that can only be written to.  Contact course vendor";
  	this.errDiagn["405"] = "The course gave an incorrect Data type.  Contact course vendor";
  
  	//member functions
  	this.LMSInitialize     = LMSInitialize;
  	this.LMSFinish         = LMSFinish;
  
  	this.LMSSetValue       = LMSSetValue;
  	this.LMSGetValue       = LMSGetValue;
  	this.LMSCommit         = LMSCommit;
  
  	this.LMSGetLastError   = LMSGetLastError;
  	this.LMSGetErrorString = LMSGetErrorString;
  	this.LMSGetDiagnostic  = LMSGetDiagnostic;
  }
  
  
  //*********************************************************
  // LMSInitialize:
  // 1. initialize the API to communicate with the SCO
  // 2. tell SCO that initialization is successful or not
  // 3. update status of lesson for student
  //*********************************************************
  function LMSInitialize(arg) {

    if (vDebug_All) alert ("initializing API...");
  
  	if (arg == ""){
  		this.APIInitialized = true;
  		this.LastErrorNum = "0";
    		//---------------------------------------------
  		// update lesson status as incomplete, since
  		// the student has now attempted the lesson,
  		// though its uncertain that they will finish
  		//---------------------------------------------
  		if (this.LMSGetValue("cmi.core.lesson_status") == "not attempted" ){
  			this.LMSSetValue("cmi.core.lesson_status","incomplete");
  		}
  
  		return "true"; //return to SCO that initialization successful
  	}
  	//---------------------------------------------------
  	// arg sent by SCO was not blank, which is an error
  	// 1. set last error as 201 - Invalid Argument Error
  	// 2. return to SCO that error occured
  	//---------------------------------------------------
  	this.LastErrorNum = "201";
  	return "false";
  }
  

  
  //*********************************************************
  // LMSFinish:
  // 1. un-initialize the API
  // 2. tell SCO that un-initialization is successful or not
  // 3. set last error to state success or an error code
  //*********************************************************
  function LMSFinish(arg) {

    if (vDebug_All) alert ("finishing API...");

  	this.LastErrorNum = "0"; //reset error code
  
  	//---------------------------------------
  	// check if API was initialized.
  	// if yes:
  	// 1. set api initialized to false
  	// 2. set last error num accordingly
  	// 3. set lesson status as complete
  	// 4. inform sco of successful finish
  	// if no:
  	// 1. set last error num accordingly
  	// 2. inform sco of unsuccessful finish
  	//---------------------------------------
  	if (this.APIInitialized) {
  		//----------------------------------
  		// check if is argument passed valid
  		//----------------------------------
  		if ((arg == "") || (arg == null)) {
  			this.APIInitialized = false;
  			this.LastErrorNum = "0";
  
  			if (this.LMSGetValue("cmi.core.lesson_status") == "incomplete" ) {
  				this.LMSSetValue("cmi.core.lesson_status","completed");
  			}
  			return "true";
  		} //end of if ((arg == "") || (arg == null))
  
  		//----------------------------------
  		// argument invalid, inform sco and
  		// set last error num accordingly
  		//----------------------------------
  		this.LastErrorNum = "201";
  		return "false";
  	} // end of if (this.APIInitialized)
  
  	//---------------------------------------
  	// API not initialized, set error and
  	// inform sco of error
  	//---------------------------------------
  	this.LastErrorNum = "301";
  	return "false";
  
  } // end of LMSFinish
  
  
  
  //*********************************************************
  // LMSSetValue:
  // 1. check if API is initialized. give error if not
  // 2. if element found in elements array, set value in
  //    values array at corresponding index
  // 3. tell SCO that set value is successful or not
  // 4. update last error num to indicate success or failure
  //*********************************************************
  function LMSSetValue(ele, val)
  { 
    if (vDebug_Key) alert ("SET API values: " + ele + " = " + val);
    
  	this.LastErrorNum = "0"; //reset error code
  	if (this.APIInitialized) {
  	  var i;

      // if posting scores then concatenate them into one value 
      if (Left(ele, 16) == "cmi.interactions") {
      	i = array_indexOf(this.elements, "cmi.interactions");
      	if (i != -1) {
      		// element found, add in values for question and answer, ignore ".type" = "choice"
          if (Right(ele, 3) == ".id") {
            // if this is not the first response then add in a pipe to separate the responses 
            if (String(this.values[i]).length == 0) {
              this.values[i] = this.values[i] + val;
            }
            else {
              this.values[i] = this.values[i] + "|" + val;
            }
          } 
          else if (Right(ele, 17) == ".student_response") {
            // if the question was not answered then replace nil by zero
            if (val == "nil") val = "0";
            // separate the question answered and the response with a colon
            this.values[i] = this.values[i] + ":" + val;
          }  

          if (vDebug_Key) alert ("SET cmi.interactions: " + this.values[i]);

      		this.LastErrorNum = "0";
      		return "true";
      	}
      }        

      // else post normal values into the array
  		i = array_indexOf(this.elements, ele);
  		if (i != -1) {
  			//element found, set value for element
  			this.values[i] = val;
  			this.LastErrorNum = "0";
  			return "true";
  		}

  		// element not implemented
  		this.LastErrorNum = "401";
  		return "false";
  	} // end of if (this.APIInitialized)
  
  	//---------------------------------------
  	// API not initialized, set error and
  	// inform sco of error
  	//---------------------------------------
  	this.LastErrorNum = "301";
  	return "false";
  }
  

  //*********************************************************
  // LMSGetValue:
  // 1. return value of element provided by calling SCO
  // 2. set appropriate error if API not initialized, element
  //    not supported or argument supplied is invalid
  //*********************************************************
  function LMSGetValue(ele) {

  	this.LastErrorNum = "0"; //reset error code
  	if (this.APIInitialized) {
  		var i = array_indexOf(this.elements,ele);
  		if (i != -1){
  			this.LastErrorNum = "0";
        if (vDebug_Key) alert ("GET API values: " + ele + " = " + this.values[i]);
  			return this.values[i];
  		}
  		this.LastErrorNum = "401"; //element not implemented
  		return "false";
  	} // end of if (this.APIInitialized)
  
  	//---------------------------------------
  	// API not initialized, set error and
  	// inform sco of error
  	//---------------------------------------
  	this.LastErrorNum = "301";
  	return "false";
  }


 
  //*********************************************************
  // LMSCommit:
  // 1. check if API is initialized. if not return error
  // 2. create form to submit exercise value to the LMS
  // 3. return appropriate error if api not initialized,
  //    invalid argument supplied, or commit not successful
  //*********************************************************
  function LMSCommit(arg){
  
    if (vDebug_Key) alert("Preparing to Commit...");
  
  	this.LastErrorNum = "0"; //reset error code
  	if (this.APIInitialized) {
  		if ((arg == "") || (arg == null)) {
  			display_values(this.elements,this.values);
  			this.LastErrorNum = "0";
  
  			//---------------------------------------------------------------------
  			// if function create_submitForm, a form which captures
  			// values passed by the exercise to API class, was successful then
  			// submit the form to transfer values to the LMS
  			//---------------------------------------------------------------------
  
    	if (create_submitForm(this.dup_elements,this.values)) {
  		header.submitForm.submit();
    	}
  
    		//---------------------------------------------------------------------
  			// form was created and submitted successfully, and data was
  			// stored in database successfully. thus return true to SCO
  			//that LMSCommit was successful.
  			//---------------------------------------------------------------------
  			return "true";
  		}
  		//invalid argument error
  		this.LastErrorNum = "201";
  		return "false";
  	} // end of if (this.APIInitialized)
  
  	//---------------------------------------
  	// API not initialized, set error and
  	// inform sco of error
  	//---------------------------------------
  	this.LastErrorNum = "301";
  	return "false";
  }
  


  //*********************************************************
  // LMSGetLastError:
  // 1. return last error encountered during SCO & API
  //    interaction. standard SCORM error number is returned.
  //*********************************************************
  function LMSGetLastError(){
  	if (this.APIInitialized) {
  		return this.LastErrorNum;
  	} // end of if (this.APIInitialized)
  
  	//---------------------------------------
  	// API not initialized, set error and
  	// inform sco of error
  	//---------------------------------------
  	this.LastErrorNum = "301";
  	return "false";
  }



  
  //*********************************************************
  // LMSGetErrorString:
  // 1. return descriptive error msg for SCORM compliant
  //    last error number set during SCO & API interaction
  //*********************************************************
  function LMSGetErrorString(errNum){
  	if (this.APIInitialized) {
  		return this.errCodes[errNum];
  	} // end of if (this.APIInitialized)
  
  	//---------------------------------------
  	// API not initialized, set error and
  	// inform sco of error
  	//---------------------------------------
  	this.LastErrorNum = "301";
  	return "false";
  }
  


  //*********************************************************
  // LMSGetDiagnostic:
  // 1. return SCORM compliant diagnostic message for
  //    error encountered during SCO & API interaction
  //*********************************************************
  function LMSGetDiagnostic(errNum){
  	if (this.APIInitialized) {
  		if (errNum == ""){
  			errNum = this.APILastError;
  		}
  		return this.errDiagn[errNum];
  	} // end of if (this.APIInitialized)
  
  	//---------------------------------------
  	// API not initialized, set error and
  	// inform sco of error
  	//---------------------------------------
  	this.LastErrorNum = "301";
  	return "false";
  }
  



  //*********************************************************
  // array_indexOf:
  // 1. returns index where the value is stored in elements
  //    array.
  //*********************************************************
  function array_indexOf(arr,val){
  	for (var i=0; i<arr.length; i++){
  		if (arr[i] == val){
  			return i;
  		}
  	}
  	return -1;
  }
  
  //*********************************************************
  // display_values:
  // 1. display values stored in elements array by exercise
  //*********************************************************
  function display_values(ele_arr,val_arr){
  	for (var i = 0; i < ele_arr.length; i++){
  	}
  }
  
  
  //*********************************************************
  // create_submitForm
  // 1. creates form that will submit exercise element values
  //    for storage in LMS database.
  //*********************************************************
  function create_submitForm(dup_ele_arr, val_arr){
  
    var submitForm = "";
    submitForm += '\n<form name="submitForm" action="/V5/ScoreModule.asp" method="post">';
    submitForm += '\n  <input type="hidden" name="vProgId"         value="' + getParameter('vProgId') + '">';
    submitForm += '\n  <input type="hidden" name="vModId"          value="' + getParameter('vModId') + '">';
    submitForm += '\n  <input type="hidden" name="vLessonLocation" value="' + val_arr[4]  + '">';
    submitForm += '\n  <input type="hidden" name="vLessonStatus"   value="' + val_arr[6]  + '">';
    submitForm += '\n  <input type="hidden" name="vScore"          value="' + val_arr[8]  + '">';
    submitForm += '\n  <input type="hidden" name="vScores"         value="' + val_arr[18] + '">';
    submitForm += '\n</form>';

    if (vDebug_Key) alert("Setting up API values for LMS...");
    if (vDebug_Key) alert(submitForm);
 
  	if (header.scorm){
  		header.scorm.innerHTML = submitForm;
  		return true;
  	}
  	return false;
  
  }

  var API = new API();
  
