<?php
use ILIAS\Filesystem\Provider\FileAccess;
use ILIAS\Filesystem\Stream\FileStream;

class ilExportWithEventoIDPlugin extends ilTestExportPlugin {
	public function getPluginName() {
		return 'ExportWithEventoID';
	}
	/**
	 * A unique identifier which describes your export type, e.g. imsm
	 * There is currently no mapping implemented concerning the filename.
	 * Feel free to create csv, xml, zip files ....
	 *
	 * @return string
	 */
	protected function getFormatIdentifier()
	{
		return 'eid';
	}
	
	/**
	 * This method should return a human readable label for your export. The string could be a translated language variable.
	 * @return string
	 */
	public function getFormatLabel()
	{
		return $this->txt('label');
	}
	
	/**
	 * This method is called if the user wants to export a test of YOUR export type
	 * If you throw an exception of type ilException with a respective language variable, ILIAS presents a translated failure message.
	 * @throws ilException
	 * @param string $export_path The path to store the export file
	 */
	protected function buildExportFile(ilTestExportFilename $export_path)
	{
		global $DIC;
		$this->test_obj = $this->getTest();
		$this->lng = $DIC->language();
		$worksheet = new ilAssExcelFormatHelper();
		$worksheet->addSheet($this->lng->txt('tst_results'));
		
		$additionalFields = $this->test_obj->getEvaluationAdditionalFields();
		
		$row = 1;
		$col = 0;
		
		if($this->test_obj->getAnonymity())
		{
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('counter'));
		}
		else
		{
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('name'));
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('login'));
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->txt('evento_id'));
		}
		
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_resultspoints'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('maximum_points'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_resultsmarks'));
		
		if($this->test_obj->getECTSOutput())
		{
			$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('ects_grade'));
		}
		
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_qworkedthrough'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_qmax'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_pworkedthrough'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_timeofwork'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_atimeofwork'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_firstvisit'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_lastvisit'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_mark_median'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_rank_participant'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_rank_median'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_total_participants'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_median'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('scored_pass'));
		$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col++) . $row, $this->lng->txt('pass'));
		
		$worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col - 1) . $row);
		
		$counter = 1;
		$data = $this->test_obj->getCompleteEvaluationData(TRUE);
		$firstrowwritten = false;
		
		foreach($data->getParticipants() as $active_id => $userdata)
		{	
			$row++;
			$col = 0;
			
			// each participant gets an own row for question column headers
			if($this->test_obj->isRandomTest())
			{
				$row++;
			}
			
			if($this->test_obj->getAnonymity())
			{
				$worksheet->setCell($row, $col++, $counter);
			}
			else
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getName());
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getLogin());
				$matriculation = ilObjUser::_getUserData([$data->getParticipant($active_id)->getUserID()])[0]['matriculation'];
				$worksheet->setCell($row, $col++, substr($matriculation, stripos($matriculation, ':')+1));
			}
			
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getReached());
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getMaxpoints());
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getMark());
			
			if($this->test_obj->getECTSOutput())
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getECTSMark());
			}
			
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getQuestionsWorkedThrough());
			$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getNumberOfQuestions());
			$worksheet->setCell($row, $col++, ($data->getParticipant($active_id)->getQuestionsWorkedThroughInPercent()) . '%');
			
			$time = $data->getParticipant($active_id)->getTimeOfWork();
			$time_seconds = $time;
			$time_hours    = floor($time_seconds/3600);
			$time_seconds -= $time_hours   * 3600;
			$time_minutes  = floor($time_seconds/60);
			$time_seconds -= $time_minutes * 60;
			$worksheet->setCell($row, $col++, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
			$time = $data->getParticipant($active_id)->getQuestionsWorkedThrough() ? $data->getParticipant($active_id)->getTimeOfWork() / $data->getParticipant($active_id)->getQuestionsWorkedThrough() : 0;
			$time_seconds = $time;
			$time_hours    = floor($time_seconds/3600);
			$time_seconds -= $time_hours   * 3600;
			$time_minutes  = floor($time_seconds/60);
			$time_seconds -= $time_minutes * 60;
			$worksheet->setCell($row, $col++, sprintf("%02d:%02d:%02d", $time_hours, $time_minutes, $time_seconds));
			$worksheet->setCell($row, $col++, new ilDateTime($data->getParticipant($active_id)->getFirstVisit(), IL_CAL_UNIX));
			$worksheet->setCell($row, $col++, new ilDateTime($data->getParticipant($active_id)->getLastVisit(), IL_CAL_UNIX));
			
			$median = $data->getStatistics()->getStatistics()->median();
			$pct = $data->getParticipant($active_id)->getMaxpoints() ? $median / $data->getParticipant($active_id)->getMaxpoints() * 100.0 : 0;
			$mark = $this->test_obj->mark_schema->getMatchingMark($pct);
			$mark_short_name = "";
			
			if(is_object($mark))
			{
				$mark_short_name = $mark->getShortName();
			}
			
			$worksheet->setCell($row, $col++, $mark_short_name);
			$worksheet->setCell($row, $col++, $data->getStatistics()->getStatistics()->rank($data->getParticipant($active_id)->getReached()));
			$worksheet->setCell($row, $col++, $data->getStatistics()->getStatistics()->rank_median());
			$worksheet->setCell($row, $col++, $data->getStatistics()->getStatistics()->count());
			$worksheet->setCell($row, $col++, $median);
			
			if($this->test_obj->getPassScoring() == SCORE_BEST_PASS)
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getBestPass() + 1);
			}
			else
			{
				$worksheet->setCell($row, $col++, $data->getParticipant($active_id)->getLastPass() + 1);
			}
			
			$startcol = $col;
			
			for($pass = 0; $pass <= $data->getParticipant($active_id)->getLastPass(); $pass++)
			{
				$col = $startcol;
				$finishdate = ilObjTest::lookupPassResultsUpdateTimestamp($active_id, $pass);
				if($finishdate > 0)
				{
					if ($pass > 0)
					{
						$row++;
						if ($this->test_obj->isRandomTest())
						{
							$row++;
						}
					}
					$worksheet->setCell($row, $col++, $pass + 1);
					if(is_object($data->getParticipant($active_id)) && is_array($data->getParticipant($active_id)->getQuestions($pass)))
					{
						$evaluatedQuestions = $data->getParticipant($active_id)->getQuestions($pass);
						
						if( $this->test_obj->getShuffleQuestions() )
						{
							// reorder questions according to general fixed sequence,
							// so participant rows can share single questions header
							$questions = array();
							foreach($this->test_obj->getQuestions() as $qId)
							{
								foreach($evaluatedQuestions as $evaledQst)
								{
									if( $evaledQst['id'] != $qId )
									{
										continue;
									}
									
									$questions[] = $evaledQst;
								}
							}
						}
						else
						{
							$questions = $evaluatedQuestions;
						}
						
						foreach($questions as $question)
						{
							$question_data = $data->getParticipant($active_id)->getPass($pass)->getAnsweredQuestionByQuestionId($question["id"]);
							$worksheet->setCell($row, $col, $question_data["reached"]);
							if($this->test_obj->isRandomTest())
							{
								// random test requires question headers for every participant
								// and we allready skipped a row for that reason ( --> row - 1)
								$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col) . ($row - 1),  preg_replace("/<.*?>/", "", $data->getQuestionTitle($question["id"])));
							}
							else
							{
								if($pass == 0 && !$firstrowwritten)
								{
									$worksheet->setFormattedExcelTitle($worksheet->getColumnCoord($col) . 1, $data->getQuestionTitle($question["id"]));
								}
							}
							$col++;
						}
						$firstrowwritten = true;
					}
				}
			}
			$counter++;
		}
		
		$excelfile = $export_path->getPathname("xlsx");
		if (!is_dir(dirname($excelfile))) {
			mkdir(dirname($excelfile), 0777, true);
		}

		$worksheet->writeToFile($excelfile);
	}
}