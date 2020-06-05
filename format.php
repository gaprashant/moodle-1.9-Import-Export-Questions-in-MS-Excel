<?php
// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * Excel format question importer.
 *
 * @package    qformat_excel
 * @copyright  2017 prashanth galagali
 * 
 */


defined('MOODLE_INTERNAL') || die();


/**
 * XLS format - a simple format for creating multiple choice questions (with
 * only one correct choice, and no feedback).
 *
 * The format looks like this:
 *
 * Question text
 * A) Choice #1
 * B) Choice #2
 * C) Choice #3
 * D) Choice #4
 * ANSWER: B
 *
 * That is,
 *  + question text all one one line.
 *  + then a number of choices, one to a line. Each line must comprise a letter,
 *    then ')' or '.', then a space, then the choice text.
 *  + Then a line of the form 'ANSWER: X' to indicate the correct answer.
 *
 * Be sure to word "All of the above" type choices like "All of these" in
 * case choices are being shuffled.
 *
 * @copyright  2017 prashanth galagali
 * 
 
[0] = qtype
[1] = multiselection / usecase
[2] = qname
[3] = qtext
[4] = OPT1
[5] = OPT2
[6] = OPT3
[7] = OPT4
[8] = OPT5
[9] = Answer
[10]= Marks
[11]= Penalty in %

 */

require_once($CFG->libdir . '/excelreader/excelreader.php');
 
class qformat_xls extends qformat_default {

    public function provide_import() {
        return true;
    }

    public function provide_export() {
        return true;
    }
	
    public function mime_type() {
        return 'application/vnd.ms-excel';
		//return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    }
	
	public function export_file_extension() {
		return '.xls';
		//return '.xlsx';
    }
	
	
    public function readquestions($questions) {
		
		//print_r(is_readable($this->filename));		
			
		$data = new Spreadsheet_Excel_Reader();

		// Set output Encoding.
		$data->setOutputEncoding('CP1251');
		
		$file = $this->filename;
		$data->read($file);
		
		$questions = array();
		
		for ($i = 2; $i <= $data->sheets[0]['numRows']; $i++) {			
				$questions[$i] = array();				
			for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
					$questions[$i][$j] = $data->sheets[0]['cells'][$i][$j];
			}
				// reset array keys starting from 0
				$questions[$i] = array_values($questions[$i]);
		}

		// reset array keys starting from 0
		$questions = array_values($questions);

		$qo = array(); 		
		
		foreach($questions as $k=>$v)
		{
				switch(strtolower(trim($v[0])))
				{	case 'multichoice' : $qo[] = $this->import_multichoice($v); break;
					case 'truefalse' : $qo[] = $this->import_truefalse($v); break;
					case 'shortanswer' : $qo[] = $this->import_shortanswer($v); break;
					case 'essay' : $qo[] = $this->import_essay($v); break;
					default : break;
				}
		}
		
		return $qo;
    }
	
	public function import_multichoice($question)
	{

		  $qo = $this->defaultquestion();
		  
		  $qo->questiontextformat = FORMAT_HTML;
		  $qo->generalfeedback = '';
		  $qo->generalfeedbackformat = FORMAT_HTML;
		 
		  $qo->fraction = array();
		  $qo->feedback = array();
		  $qo->correctfeedback = $this->text_field('');
		  $qo->partiallycorrectfeedback = $this->text_field('');
		  $qo->incorrectfeedback = $this->text_field('');
		  
		  $qo->qtype = 'multichoice';
		  
		  $qo->single = ($question[1]) ? 0 : 1; // if multiselection is true assign 0;
		  
		  $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
		  
		  $qo->questiontext = htmlspecialchars(trim($question[3]), ENT_NOQUOTES);
		  $qo->questiontext = '<pre>'.$qo->questiontext.'</pre>';
		  

		 // $qo->questiontext = str_replace('&lt;br&gt;','<br>',$qo->questiontext);
		 // $qo->questiontext = str_replace('&lt;br /&gt;','<br />',$qo->questiontext);		  
		  
		  
		  $qo->answer = array();
							  
		  // ---------- list of options
		  $qo->answer[] = $this->text_field(htmlspecialchars(trim($question[4]), ENT_NOQUOTES));
		  $qo->fraction[] = 0;
		  $qo->feedback[] = $this->text_field('');
		  $qo->answer[] = $this->text_field(htmlspecialchars(trim($question[5]), ENT_NOQUOTES));
		  $qo->fraction[] = 0;
		  $qo->feedback[] = $this->text_field('');
		  $qo->answer[] = $this->text_field(htmlspecialchars(trim($question[6]), ENT_NOQUOTES));
		  $qo->fraction[] = 0;
		  $qo->feedback[] = $this->text_field('');
		  $qo->answer[] = $this->text_field(htmlspecialchars(trim($question[7]), ENT_NOQUOTES));
		  $qo->fraction[] = 0;
		  $qo->feedback[] = $this->text_field('');
		  $qo->answer[] = $this->text_field(htmlspecialchars(trim($question[8]), ENT_NOQUOTES));
		  $qo->fraction[] = 0;
		  $qo->feedback[] = $this->text_field('');					
		  
		  
		  $key = $question[9];
		  $key = (strpos($key,','))? explode(',', $key) : $key;		  
		  
		  // true answer for single and multi selections

		  $kans = floatval(1/sizeof($key));

		  if(gettype($key) == 'array')
		  {
				foreach($key as $k => $v)
				{	$kv = 	filter_var($v, FILTER_SANITIZE_NUMBER_INT);
					$qo->fraction[$kv - 1] = $kans; //1;
				}
		  }
		  else if(gettype($key) == 'string')
		  {
			  $key = filter_var($question[9], FILTER_SANITIZE_NUMBER_INT);			  
			  $qo->fraction[$key-1] = 1;
		  }
		  
		  $qo->defaultmark = 	$question[10];
		  $qo->penalty = 	$question[11];
		  
		  return $qo;
	}
	
	
	public function import_truefalse($question)
	{
		  $qo = $this->defaultquestion();
		  
		  $qo->questiontextformat = FORMAT_HTML;
		  $qo->generalfeedback = '';
		  $qo->generalfeedbackformat = FORMAT_HTML;

		   
		  $qo->qtype = 'truefalse';
		  
		  $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
		  
		  $qo->questiontext = htmlspecialchars(trim($question[3]), ENT_NOQUOTES);
		  $qo->questiontext = '<pre>'.$qo->questiontext.'</pre>';		  
		  
		  // answer
		  $key = filter_var($question[9], FILTER_SANITIZE_NUMBER_INT);
		  $key = intval($key-1); // array starts from 0;
		  
		  if($key == 0)
		  {	$qo->answer = (trim($question[4])) ? 1 : '';	}
		  else if($key == 1)
		  {	$qo->answer = (trim($question[5])) ? 1 : '';	}

			
		  $qo->correctanswer = $qo->answer;
			
		  $qo->feedbackfalse = $this->text_field('');
		  $qo->feedbacktrue = $this->text_field('');
		  
		  
		  $qo->defaultmark = 	$question[10];
		  // $qo->penalty = 	$question[11]; // there is no penalty for true false answer
		  
		  return $qo;
		
		}
		
		
	public function import_shortanswer($question)
	{		
		$qo = $this->defaultquestion();
		$qo->questiontextformat = FORMAT_HTML;
		$qo->generalfeedback = '';
		$qo->generalfeedbackformat = FORMAT_HTML;
		
		  
	    $qo->qtype = 'shortanswer';
		$qo->usecase = ($question[1]) ? 1 : 0; // Use case
		 
		$qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
		  
		$qo->questiontext = htmlspecialchars(trim($question[3]), ENT_NOQUOTES);
		$qo->questiontext = '<pre>'.$qo->questiontext.'</pre>';		
		
		
		$qo->answer = array();
		$qo->fraction = array();
		$qo->feedback = array();
					  
		// There will be only one correct answer. and fraction is also first only 
		// No need to check answer field of excel sheet always first option will be correct
		$qo->answer[0] = htmlspecialchars(trim($question[4]), ENT_NOQUOTES);
		$qo->fraction[0] = 1;
		$qo->feedback[0] = $this->text_field('');
		
		$qo->defaultmark = 	$question[10];
		$qo->penalty = 	$question[11];
	
		return $qo;
	}
	
	public function import_essay($question)
	{
		$qo = $this->defaultquestion();
		$qo->questiontextformat = FORMAT_HTML;
		$qo->generalfeedback = '';
		$qo->generalfeedbackformat = FORMAT_HTML;
		
	    $qo->qtype = 'essay';
		 
		$qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
		  
		$qo->questiontext = htmlspecialchars(trim($question[3]), ENT_NOQUOTES);
		$qo->questiontext = '<pre>'.$qo->questiontext.'</pre>';		  
		  		  
		$qo->defaultmark = 	$question[10];
		//$qo->penalty = 	$question[11]; // penalty not required
								
								
		$qo->responseformat = 'editor';
		$qo->responsefieldlines = 15;
		$qo->responserequired = 1;
		$qo->attachments = 0;
		$qo->attachmentsrequired = 0;
		
		$qo->graderinfo = $this->text_field('');
		$qo->responsetemplate =  $this->text_field('');
				
		return $qo;
	}
	
	
	public function readquestion($lines) {
        // This is no longer needed but might still be called by default.php.
        return;
    }
	
	protected function text_field($text) {
        return array(
            'text' => htmlspecialchars(trim($text), ENT_NOQUOTES),
            'format' => FORMAT_HTML,
            //'files' => array(),
        );
    }
	
	
	protected function presave_process($content) {
        // Override to allow us to add xml headers and footers.
		
		$strr = array();
		$strr[] = '<pre>';
		$strr[] = '</pre>';
		
		$content = str_replace($strr,'',$content);
		
		$sep = "\t";
        return 'QType'.$sep.'Multichoice/Usecase'.$sep.'QName'.$sep.'QText'.$sep.'OPT1'.$sep.'OPT2'.$sep.'OPT3'.$sep.'OPT4'.$sep.'OPT5'.$sep.'Answer'.$sep.'Default Mark'.$sep.'Penalty'.$content;
		
    }
	
	
	protected function writequestion($question) {
        global $OUTPUT;
		
		if (($question->qtype=='multichoice') || ($question->qtype=='truefalse') || ($question->qtype=='shortanswer') || ($question->qtype=='essay')) 
		{	
			$expout = '';
			$sep = "\t";
			
			switch($question->qtype)
			{	case 'multichoice' : $expout .= $this->export_multichoice($question); break;
				case 'truefalse' : $expout .= $this->export_truefalse($question); break;
				case 'shortanswer' : $expout .= $this->export_shortanswer($question); break;
				case 'essay' : $expout .= $this->export_essay($question); break;
				default : break;
			}		
        	return $expout;
		}
		else
		{	return false;	}
		
		return false;	
    }
	
	
	public function export_multichoice($question)
	{
		$sep = "\t";
		$expout = "";
		$expout .= "multichoice".$sep;						// question type
		$expout .= (($question->options->single)? 'FALSE':'TRUE').$sep;						// multiselect / use case
		$expout .= $question->name.$sep;				// question name
		$expout .= $question->questiontext.$sep;		// question text
	

		$opt = 0;
		$ans = '';
		
		// list of options
		foreach($question->options->answers as $k=>$v)
		{	$opt++;
			$expout .= $v->answer.$sep;	
			if($v->fraction > 0)	// valuce can be 0.5
			{	$ans .= 'OPT'.$opt.',';	}
		}	
		
		$ans = substr($ans, 0, -1); // remove last ,
		
		//  if options less than 5 - hardcoded
		if(sizeof($question->options->answers) < 5)
		{
			for($i = sizeof($question->options->answers); $i<5; $i++)
			{	$expout .= "".$sep;  // extra or empty options
			}
		}
		
		$expout .= $ans.$sep;						// answer
		$expout .= $question->defaultmark.$sep;		// default mark field
		$expout .= $question->penalty.$sep;			// penalty field
		//$expout .= '\n';

		return $expout;
	}
	
	public function export_shortanswer($question)
	{		
		$sep = "\t";
		$expout = "";
		
		$expout .= "shortanswer".$sep;						// question type
		$expout .= (($question->options->usecase)? 'TRUE':'FALSE').$sep;	// multiselect / use case
		$expout .= $question->name.$sep;				// question name
		$expout .= $question->questiontext.$sep;		// question text
		
		// options 1 only
		foreach($question->options->answers as $k=>$v)
		{	$expout .= $v->answer.$sep;		}
		
		
		$expout .= "".$sep;							// option 2
		$expout .= "".$sep;							// option 3
		$expout .= "".$sep;							// option 4
		$expout .= "".$sep;							// option 5
		$expout .= "OPT1".$sep;							// answer // hard coded
		$expout .= $question->defaultmark.$sep;		// default mark field
		$expout .= $question->penalty.$sep;							// penalty field
		//$expout .= '\n';
		
		return $expout;
	}
	
	
	public function export_truefalse($question)
	{
		$sep = "\t";
		$expout = "";
		$expout .= "truefalse".$sep;						// question type
		$expout .= "".$sep;							// multiselect / use case
		$expout .= $question->name.$sep;				// question name
		$expout .= $question->questiontext.$sep;		// question text
	
		$arr = array();
		$opt = 0;
		$ans = '';
		
		// options 1 and 2
		foreach($question->options->answers as $k=>$v)
		{	$opt++;
			$expout .= $v->answer.$sep;	
			if(intval($v->fraction))
			{	$ans = 'OPT'.$opt;	}
		}
		
		if(sizeof($question->options->answers) == 1)	// there 2 fields if in case if only one field entered
		{	$expout .= "".$sep; }
		
		$expout .= "".$sep;							// option 3
		$expout .= "".$sep;							// option 4
		$expout .= "".$sep;							// option 5
		$expout .= $ans.$sep;							// answer
		$expout .= $question->defaultmark.$sep;		// default mark field
		$expout .= "".$sep;							// penalty field
		//$expout .= '\n';
		
		return $expout;
	}
		
	public function export_essay($question)
	{
		
		// excel columns
		/*
		[0] = qtype
		[1] = multiselection / usecase
		[2] = qname
		[3] = qtext
		[4] = OPT1
		[5] = OPT2
		[6] = OPT3
		[7] = OPT4
		[8] = OPT5
		[9] = Answer
		[10]= Marks
		[11]= Penalty in %
		*/
		
		$sep = "\t";
		$expout = "";
		$expout .= "essay".$sep;						// question type
		$expout .= "".$sep;							// multiselect / use case
		$expout .= $question->name.$sep;				// question name
		$expout .= $question->questiontext.$sep;		// question text
		$expout .= "".$sep;							// option 1
		$expout .= "".$sep;							// option 2
		$expout .= "".$sep;							// option 3
		$expout .= "".$sep;							// option 4
		$expout .= "".$sep;							// option 5
		$expout .= "".$sep;							// answer
		$expout .= $question->defaultmark.$sep;		// default mark field
		$expout .= "".$sep;							// penalty field
		//$expout .= '\n';		
		
		return $expout;
	}
	
	
	
	/**
     * Do the export
     * For most types this should not need to be overrided
     * @return stored_file
     */
   /* public function exportprocess() {
        global $CFG, $OUTPUT, $DB, $USER;

        // get the questions (from database) in this category
        // only get q's with no parents (no cloze subquestions specifically)
        if ($this->category) {
            $questions = get_questions_category($this->category, true);
        } else {
            $questions = $this->questions;
        }

        $count = 0;
		
		print_r($questions);
		die('ok');

        // results are first written into string (and then to a file)
        // so create/initialize the string here
        $expout = "";

        // track which category questions are in
        // if it changes we will record the category change in the output
        // file if selected. 0 means that it will get printed before the 1st question
        $trackcategory = 0;

        // iterate through questions
        foreach ($questions as $question) {
            // used by file api
            $contextid = $DB->get_field('question_categories', 'contextid',
                    array('id' => $question->category));
            $question->contextid = $contextid;

            // do not export hidden questions
            if (!empty($question->hidden)) {
                continue;
            }

            // do not export random questions
            if ($question->qtype == 'random') {
                continue;
            }

            // check if we need to record category change
            if ($this->cattofile) {
                if ($question->category != $trackcategory) {
                    $trackcategory = $question->category;
                    $categoryname = $this->get_category_path($trackcategory, $this->contexttofile);

                    // create 'dummy' question for category export
                    $dummyquestion = new stdClass();
                    $dummyquestion->qtype = 'category';
                    $dummyquestion->category = $categoryname;
                    $dummyquestion->name = 'Switch category to ' . $categoryname;
                    $dummyquestion->id = 0;
                    $dummyquestion->questiontextformat = '';
                    $dummyquestion->contextid = 0;
                    $expout .= $this->writequestion($dummyquestion) . "\n";
                }
            }

            // export the question displaying message
            $count++;

            if (question_has_capability_on($question, 'view', $question->category)) {
                $expout .= $this->writequestion($question, $contextid) . "\n";
            }
        }

        // continue path for following error checks
        $course = $this->course;
        $continuepath = "{$CFG->wwwroot}/question/export.php?courseid={$course->id}";

        // did we actually process anything
        if ($count==0) {
            print_error('noquestions', 'question', $continuepath);
        }

        // final pre-process on exported data
        $expout = $this->presave_process($expout);
        return $expout;
    }
	*/
}


