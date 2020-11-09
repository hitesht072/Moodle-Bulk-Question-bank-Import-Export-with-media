<?php

/**
 * Excel format question importer.
 *
 * @package    qformat_Excel
 * @copyright  2020 Hitesh Kumar
 * @license    
 */
defined('MOODLE_INTERNAL') || die();

/**
 * XLS format - a simple format for creating multiple choice questions (with
 */
require_once($CFG->dirroot . '/question/format/xls/excelreader/excelreader.php');

class qformat_xls extends qformat_default {

    protected $_images;

    /** @var string path to the temporary directory. */
    public $tempdir = '';

    public function provide_import() {
        return true;
    }

    public function provide_export() {
        return false;
    }

    public function mime_type() {
        return mimeinfo('type', '.zip');
    }

    protected function escapedchar_post($string) {
        // Replaces placeholders with corresponding character AFTER processing is done.
        $placeholders = array("&&058;", "&&035;", "&&061;", "&&123;", "&&125;", "&&126;", "&&010");
        $characters = array(":", "#", "=", "{", "}", "~", "\n");
        $string = str_replace($placeholders, $characters, $string);
        return $string;
    }

    public function importpostprocess() {
        global $CFG, $DB;

        if ($this->tempdir != '') {
            fulldelete($this->tempdir);
        }
        return true;
    }

    /**
     * Store an image file in a draft filearea
     * @param array $text, if itemid element doesn't exist it will be created
     * @param string $tempdir path to root of image tree
     * @param string $filepathinsidetempdir path to image in the tree
     * @param string $filename image's name
     * @return string new name of the image as it was stored
     */
    protected function store_file_for_text_field(&$text, $tempdir, $filepathinsidetempdir, $filename) {
        global $USER;

        $fs = get_file_storage();
        if (empty($text['itemid'])) {
            $text['itemid'] = file_get_unused_draft_itemid();
        }
        // As question file areas don't support subdirs,
        // convert path to filename.
        // So that medias with same name can be imported.
        if ($filepathinsidetempdir == '.') {
            $newfilename = clean_param($filename, PARAM_FILE);
        } else {
            $newfilename = clean_param(str_replace('/', '__', $filepathinsidetempdir . '__' . $filename), PARAM_FILE);
        }
        $filerecord = array(
            'contextid' => context_user::instance($USER->id)->id,
            'component' => 'user',
            'filearea' => 'draft',
            'itemid' => $text['itemid'],
            'filepath' => '/',
            'filename' => $newfilename,
        );
        if ($filepathinsidetempdir == '.') {
            $fs->create_file_from_pathname($filerecord, $tempdir . '/' . $filename);
        } else {
            $fs->create_file_from_pathname($filerecord, $tempdir . '/' . $filepathinsidetempdir . '/' . $filename);
        }
        return $newfilename;
    }

    public function export_file_extension() {
        return '.zip';
    }

    /**
     * Parse the text
     * @param string $text the text to parse
     * @param integer $defaultformat text format
     * @return array with keys text, format, itemid.
     *
     */
    protected function parse_text_with_format($text, $defaultformat = FORMAT_MOODLE) {
        // Parameter defaultformat is ignored we set format to be html in all cases.
        return $this->text_field(trim($this->escapedchar_post($text)));
    }

    /**
     * Return content of all files containing questions,
     * as an array one element for each file found,
     * For each file, the corresponding element is an array of lines.
     * @param string $filename name of file
     * @return mixed contents array or false on failure
     */
    public function readdata($filename) {
        $uniquecode = time();
        $this->tempdir = make_temp_directory('xls/' . $uniquecode);

        if (file_exists($filename)) {
            if (!copy($filename, $this->tempdir . '/xls.zip')) {
                $this->error(get_string('cannotcopybackup', 'question'));
                fulldelete($this->tempdir);
                return false;
            }
            $packer = get_file_packer('application/zip');
            if ($packer->extract_to_pathname($this->tempdir . '/xls.zip', $this->tempdir)) {
                // Search for a text file in the zip archive.
                // TODO ? search it, even if it is not a root level ?
                $filenames = array();
                $iterator = new DirectoryIterator($this->tempdir);
                foreach ($iterator as $fileinfo) {
                    if ($fileinfo->isFile() && strtolower(pathinfo($fileinfo->getFilename(), PATHINFO_EXTENSION)) == 'xls') {
                        $filenames[] = $fileinfo->getFilename();
                    }
                }
                if ($filenames) {
                    $this->filename = $this->tempdir . '/' . $filenames[0];
                    return true;
                } else {
                    $this->error(get_string('noxlsfile', 'xls'));
                    fulldelete($this->temp_dir);
                }
            } else {
                $this->error(get_string('cannotunzip', 'question'));
                fulldelete($this->temp_dir);
            }
        } else {
            $this->error(get_string('cannotreaduploadfile', 'error'));
            fulldelete($this->tempdir);
        }
        return false;
    }

    public function readquestions($questions) {
        global $CFG;
        $data = new Spreadsheet_Excel_Reader();

        // Set output Encoding.
//        $data->setOutputEncoding('CP1251');
        $data->setOutputEncoding('UTF-8');

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

        foreach ($questions as $k => $v) {
            if (empty($v[1]) || empty($v[2]) || empty($v[3])) { // qtype, qname, qtext
                continue;
            }
            switch (strtolower(trim($v[1]))) {
                case 'multichoice' :
                    $qo[] = $this->import_multichoice($v);
                    break;
                case 'truefalse' :
                    $qo[] = $this->import_truefalse($v);
                    break;
                case 'shortanswer' :
                    $qo[] = $this->import_shortanswer($v);
                    break;
                case 'essay' :
                    $qo[] = $this->import_essay($v);
                    break;
                case 'gapselect' :
                    $qo[] = $this->import_gapselect($v);
                    break;
                case 'ddwtos' :
                    $qo[] = $this->import_ddwtos($v);
                    break;
                case 'match' :
                    $qo[] = $this->import_match($v);
                    break;
                default : break;
            }
        }

        $qo = array_filter($qo); // remove empty values or false values
        $qo = array_values($qo); // reset array keys starting from 0

        return $qo;
    }


    private function _update_tags($question) {
        $tags[] = $question[7];
        $tags[] = $question[8];
        return $tags;
    }

    public function import_multichoice($question) {

        // if options are blank
        if (empty($question[4])) {
            return false;
        }

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

        $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
        if (empty($qo->name)) {
            $qo->name = $this->create_default_question_name($question[2], get_string('questionname', 'question'));
        }

        // Get questiontext format from questiontext.
        $text = $this->parse_text_with_format($question[3]);
        $qo->questiontextformat = $text['format'];
        $qo->questiontext = $text['text'];
        if (!empty($text['itemid'])) {
            $qo->questiontextitemid = $text['itemid'];
        }

        // ---------- list of options
        $answers = explode('|', $question[4]);
        $i = 1;
        foreach ($answers as $sa) {
            $qo->answer[] = $this->text_field($sa);
            $qo->fraction[] = 0;
            $qo->feedback[] = $this->text_field('');
            $i++;
        }

        $key = $question[5];
        $key = (strpos($key, ',')) ? explode(',', $key) : $key;

        $kans = floatval(1 / sizeof($key));

        if (gettype($key) == 'array') {
            foreach ($key as $k => $v) {
                $kv = filter_var($v, FILTER_SANITIZE_NUMBER_INT);
                $qo->fraction[$kv - 1] = $kans; //1;
            }

            $qo->single = 0; // if multiselection is true assign 0;
        } else if (gettype($key) == 'string') {
            $key = filter_var($question[5], FILTER_SANITIZE_NUMBER_INT);
            $qo->fraction[$key - 1] = 1;
            $qo->single = 1; // if multiselection is true assign 0;
        }

        $qo->defaultmark = (!empty($question[6])) ? $question[6] : 1; // default value hardcoded 
        $qo->penalty = 0.33; // default value hardcoded 
        
        //update tags fileds
        $qo->tags = $this->_update_tags($question);
        $qo->extras = $this->_update_extra($question);

        return $qo;
    }

    public function import_gapselect($question) {
        // if answer field is blank, skip the question from loop
        if (empty($question[4])) {
            return false;
        }

        $qo = $this->defaultquestion();

        $qo->questiontextformat = FORMAT_HTML;
        $qo->generalfeedback = '';
        $qo->generalfeedbackformat = FORMAT_HTML;

        $qo->fraction = array();
        $qo->feedback = array();
        $qo->correctfeedback = $this->text_field('');
        $qo->partiallycorrectfeedback = $this->text_field('');
        $qo->incorrectfeedback = $this->text_field('');

        $qo->qtype = 'gapselect';

        $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
        if (empty($qo->name)) {
            $qo->name = $this->create_default_question_name($question[2], get_string('questionname', 'question'));
        }

        // Get questiontext format from questiontext.
        $text = $this->parse_text_with_format($question[3]);
        $qo->questiontextformat = $text['format'];
        $qo->questiontext = $text['text'];
        if (!empty($text['itemid'])) {
            $qo->questiontextitemid = $text['itemid'];
        }

        $qo->answer = array();

        $answers = explode('|', $question[4]);
        $qo->noanswers = count($answers);
        foreach ($answers as $sa) {
            $qoanswer['answer'] = $sa;
            $qoanswer['choicegroup'] = 1; //choce group , currenty only in one group
            $qo->choices[] = $qoanswer;
        }

        $qo->defaultmark = (!empty($question[6])) ? $question[6] : 1; // default value hardcoded 
        $qo->penalty = 0.33; // default value hardcoded 
        
        //update tags fileds
        $qo->tags = $this->_update_tags($question);
        $qo->extras = $this->_update_extra($question);

        return $qo;
    }

    public function import_ddwtos($question) {
        // if answer field is blank, skip the question from loop
        if (empty($question[4])) {
            return false;
        }

        $qo = $this->defaultquestion();

        $qo->questiontextformat = FORMAT_HTML;
        $qo->generalfeedback = '';
        $qo->generalfeedbackformat = FORMAT_HTML;

        $qo->fraction = array();
        $qo->feedback = array();
        $qo->correctfeedback = $this->text_field('');
        $qo->partiallycorrectfeedback = $this->text_field('');
        $qo->incorrectfeedback = $this->text_field('');

        $qo->qtype = 'ddwtos';
        $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
        if (empty($qo->name)) {
            $qo->name = $this->create_default_question_name($question[2], get_string('questionname', 'question'));
        }

        // Get questiontext format from questiontext.
        $text = $this->parse_text_with_format($question[3]);
        $qo->questiontextformat = $text['format'];
        $qo->questiontext = $text['text'];
        if (!empty($text['itemid'])) {
            $qo->questiontextitemid = $text['itemid'];
        }

        $qo->answer = array();

        $answers = explode('|', $question[4]);
        $qo->noanswers = count($answers);
        foreach ($answers as $sa) {
            $qoanswer['answer'] = $sa;
            $qoanswer['choicegroup'] = 1; //choce group , currenty only in one group
            $qo->choices[] = $qoanswer;
        }

        $qo->defaultmark = (!empty($question[6])) ? $question[6] : 1; // default value hardcoded 
        $qo->penalty =  0.33; // default value hardcoded 
        
        //update tags fileds
        $qo->tags = $this->_update_tags($question);
        $qo->extras = $this->_update_extra($question);

        return $qo;
    }

    public function import_match($question) {
        // if answer field is blank, skip the question from loop
        if (empty($question[4])) {
            return false;
        }

        $qo = $this->defaultquestion();

        $qo->questiontextformat = FORMAT_HTML;
        $qo->generalfeedback = '';
        $qo->generalfeedbackformat = FORMAT_HTML;

        $qo->fraction = array();
        $qo->feedback = array();
        $qo->correctfeedback = $this->text_field('');
        $qo->partiallycorrectfeedback = $this->text_field('');
        $qo->incorrectfeedback = $this->text_field('');

        $qo->qtype = 'match';

        $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
        if (empty($qo->name)) {
            $qo->name = $this->create_default_question_name($question[2], get_string('questionname', 'question'));
        }

        // Get questiontext format from questiontext.
        $text = $this->parse_text_with_format($question[3]);
        $qo->questiontextformat = $text['format'];
        $qo->questiontext = $text['text'];
        if (!empty($text['itemid'])) {
            $qo->questiontextitemid = $text['itemid'];
        }

        $qo->subquestions = array();

        $answers = explode('|', $question[4]);
        $qo->noanswers = count($answers);
        foreach ($answers as $sa) {
            $question_answer = explode('->', $sa);
            $qo->subanswers[] = $question_answer[1];
            $qo->subquestions[] = $this->parse_text_with_format($question_answer[0]);
        }

        $qo->defaultmark = (!empty($question[6])) ? $question[6] : 1; // default value hardcoded 
        $qo->penalty = 0.33; // default value hardcoded 
        
        //update tags fileds
        $qo->tags = $this->_update_tags($question);
        $qo->extras = $this->_update_extra($question);

        return $qo;
    }

    public function import_truefalse($question) {

        // if options are blank
        if (empty($question[4])) {
            return false;
        }

        $qo = $this->defaultquestion();

        $qo->questiontextformat = FORMAT_HTML;
        $qo->generalfeedback = '';
        $qo->generalfeedbackformat = FORMAT_HTML;

        $qo->qtype = 'truefalse';

        $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
        if (empty($qo->name)) {
            $qo->name = $this->create_default_question_name($question[2], get_string('questionname', 'question'));
        }

        // Get questiontext format from questiontext.
        $text = $this->parse_text_with_format($question[3]);
        $qo->questiontextformat = $text['format'];
        $qo->questiontext = $text['text'];
        if (!empty($text['itemid'])) {
            $qo->questiontextitemid = $text['itemid'];
        }
        // answer
        $key = filter_var($question[5], FILTER_SANITIZE_NUMBER_INT);
        $key = intval($key - 1); // array starts from 0;

        $answers = explode('|', $question[4]);
        if(trim(strtolower($answers[$key])) == 'true'){
            $qo->answer = 1;
        }else{
            $qo->answer = 0;
        }
        
        $qo->correctanswer = $qo->answer;

        $qo->feedbackfalse = $this->text_field('');
        $qo->feedbacktrue = $this->text_field('');


        $qo->defaultmark = (!empty($question[6])) ? $question[6] : 1; // default value hardcoded 
        
        //update tags fileds
        $qo->tags = $this->_update_tags($question);
        $qo->extras = $this->_update_extra($question);

        return $qo;
    }

    public function import_shortanswer($question) {

        // if options are blank
        if (empty($question[4])) {
            return false;
        }

        $qo = $this->defaultquestion();
        $qo->questiontextformat = FORMAT_HTML;
        $qo->generalfeedback = '';
        $qo->generalfeedbackformat = FORMAT_HTML;


        $qo->qtype = 'shortanswer';
//        $qo->usecase = ($question[1]) ? 1 : 0; // Use case
        $qo->usecase = 0; // Use case

        $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
        if (empty($qo->name)) {
            $qo->name = $this->create_default_question_name($question[2], get_string('questionname', 'question'));
        }

        // Get questiontext format from questiontext.
        $text = $this->parse_text_with_format($question[3]);
        $qo->questiontextformat = $text['format'];
        $qo->questiontext = $text['text'];
        if (!empty($text['itemid'])) {
            $qo->questiontextitemid = $text['itemid'];
        }
        $qo->answer = array();
        $qo->fraction = array();
        $qo->feedback = array();

        // There will be only one correct answer. and fraction is also first only 
        // No need to check answer field of excel sheet always first option will be correct
        $qo->answer[0] = htmlspecialchars(trim($question[5]), ENT_NOQUOTES);
        $qo->fraction[0] = 1;
        $qo->feedback[0] = $this->text_field('');

        $qo->defaultmark = (!empty($question[6])) ? $question[6] : 1; // default value hardcoded 		
        $qo->penalty =  0.33; // default value hardcoded 
         
        //update tags fileds
        $qo->tags = $this->_update_tags($question);
        $qo->extras = $this->_update_extra($question);

        return $qo;
    }

    public function import_essay($question) {
        $qo = $this->defaultquestion();
        $qo->questiontextformat = FORMAT_HTML;
        $qo->generalfeedback = '';
        $qo->generalfeedbackformat = FORMAT_HTML;

        $qo->qtype = 'essay';

        $qo->name = htmlspecialchars(trim($question[2]), ENT_NOQUOTES);
        if (empty($qo->name)) {
            $qo->name = $this->create_default_question_name($question[2], get_string('questionname', 'question'));
        }

        // Get questiontext format from questiontext.
        $text = $this->parse_text_with_format($question[3]);
        $qo->questiontextformat = $text['format'];
        $qo->questiontext = $text['text'];
        if (!empty($text['itemid'])) {
            $qo->questiontextitemid = $text['itemid'];
        }

        $qo->defaultmark = (!empty($question[6])) ? $question[6] : 1; // default value hardcoded 		
        //$qo->penalty = 	$question[11]; // penalty not required

        $qo->responseformat = 'editor';
        $qo->responsefieldlines = 15;
        $qo->responserequired = 1;
        $qo->attachments = 0;
        $qo->attachmentsrequired = 0;

        $qo->graderinfo = $this->text_field('');
        $qo->responsetemplate = $this->text_field('');

        //update tags fileds
        $qo->tags = $this->_update_tags($question);
        $qo->extras = $this->_update_extra($question);

        return $qo;
    }

    public function readquestion($lines) {
        // This is no longer needed but might still be called by default.php.
        return;
    }

    public function text_field($text) {
        if (empty($text)) {
            return '';
        }
        $data = array();

        preg_match_all('|"@@PLUGINFILE@@/([^"]*)"|i', $text, $out); // Find all pluginfile refs.
        $filepaths = array();

        foreach ($out[1] as $path) {
            $fullpath = $this->tempdir . '/' . $path;

            if (is_readable($fullpath) && !in_array($path, $filepaths)) {
                $dirpath = dirname($path);
                $filename = basename($path);
                $newfilename = $this->store_file_for_text_field($data, $this->tempdir, $dirpath, $filename);
                $text = preg_replace("|@@PLUGINFILE@@/$path|", "@@PLUGINFILE@@/" . $newfilename, $text);
                $filepaths[] = $path;
            }
        }
//        $escaped = preg_replace(array('#”#u', '#\\\\#u', '#[”"]#u'), array('"', '\\\\\\\\', '\"'), $text);    
        $escaped = utf8_encode($text);
        $data['text'] = $escaped;
        $data['format'] = FORMAT_HTML;

        return $data;
    }

    protected function presave_process($content) {
        // Override to allow us to add xml headers and footers.

        $strr = array();
        $strr[] = '<pre>';
        $strr[] = '</pre>';

        $content = str_replace($strr, '', $content);

        $sep = "\t";
        return 'QType' . $sep . 'Multichoice/Usecase' . $sep . 'QName' . $sep . 'QText' . $sep . 'OPT1' . $sep . 'OPT2' . $sep . 'OPT3' . $sep . 'OPT4' . $sep . 'OPT5' . $sep . 'Answer' . $sep . 'Default Mark' . $sep . 'Penalty' . $content;
    }

}
