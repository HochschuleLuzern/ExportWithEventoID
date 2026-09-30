<?php

use ILIAS\Test\ExportImport\ExportFilename;

class ilExportWithEventoIDPlugin extends ilTestExportPlugin {

	/**
	 * Formerly EXCEL_BACKGROUND_COLOR from Modules/Test/classes/inc.AssessmentConstants.php,
	 * which no longer exists in ILIAS 10. Same value as
	 * ILIAS\Test\ExportImport\ResultsExportExcel::EXCEL_BACKGROUND_COLOR.
	 */
	private const EXCEL_BACKGROUND_COLOR = 'C0C0C0';

	protected ilObjTest $test_obj;
	protected ilLanguage $lng;

	public function getPluginName() : string {
		return 'ExportWithEventoID';
	}
	/**
	 * A unique identifier which describes your export type, e.g. imsm
	 * There is currently no mapping implemented concerning the filename.
	 * Feel free to create csv, xml, zip files ....
	 *
	 * @return string
	 */
	protected function getFormatIdentifier() : string
	{
		return 'eid';
	}
	
	/**
	 * This method should return a human readable label for your export. The string could be a translated language variable.
	 * @return string
	 */
	public function getFormatLabel() : string
	{
		return $this->txt('label');
	}
	
	/**
	 * This method is called if the user wants to export a test of YOUR export type
	 * If you throw an exception of type ilException with a respective language variable, ILIAS presents a translated failure message.
	 * @throws ilException
	 * @param ExportFilename $export_path The path to store the export file
	 */
	protected function buildExportFile(ExportFilename $export_path): void
	{
		global $DIC;
		$this->test_obj = $this->getTest();
		$this->lng = $DIC->language();
		// ilAssExcelFormatHelper was removed in ILIAS 10; plain ilExcel already writes
		// non numeric values as explicit strings and strips tags.
		$worksheet = new ilExcel();
		$worksheet->addSheet($this->lng->txt('tst_results'));
		
		$row = 1;
		$col = 0;
		
		if($this->test_obj->getAnonymity())
		{
			$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('counter'));
		}
		else
		{
			$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('name'));
			$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('login'));
			$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->txt('evento_id'));
		}
		
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_resultspoints'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('maximum_points'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_resultsmarks'));

        /*
        ECTS support dropped by ILIAS
		if($this->test_obj->getECTSOutput())
		{
			$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('ects_grade'));
		}
		*/
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_qworkedthrough'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_qmax'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_pworkedthrough'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_timeontask'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_atimeofwork'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_firstvisit'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_lastvisit'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_mark_median'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_rank_participant'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_rank_median'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_total_participants'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('tst_stat_result_median'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('scored_pass'));
		$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col++) . $row, $this->lng->txt('pass'));
		
		$worksheet->setBold('A' . $row . ':' . $worksheet->getColumnCoord($col - 1) . $row);
		
		$counter = 1;
		$data = $this->test_obj->getCompleteEvaluationData(ilTestEvaluationData::FILTER_BY_NONE, '');
		$statistics = $data->getStatistics();
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
				$worksheet->setCell($row, $col++, $userdata->getName());
				$worksheet->setCell($row, $col++, $userdata->getLogin());
				$matriculation = ilObjUser::_getUserData([$userdata->getUserID()])[0]['matriculation'] ?? '';
				$worksheet->setCell($row, $col++, substr($matriculation, stripos($matriculation, ':') + 1));
			}
			
			$worksheet->setCell($row, $col++, $userdata->getReached());
			$worksheet->setCell($row, $col++, $userdata->getMaxpoints());
			// getMark() returns an ILIAS\Test\Scoring\Marks\Mark object since ILIAS 10
			$worksheet->setCell($row, $col++, $this->getMarkShortName($userdata));
			
            /*
            ECTS support dropped by ILIAS
        	if($this->test_obj->getECTSOutput())
			{
				$worksheet->setCell($row, $col++, $userdata->getECTSMark());
			}
            */
			
			$worksheet->setCell($row, $col++, $userdata->getQuestionsWorkedThrough());
			$worksheet->setCell($row, $col++, $userdata->getNumberOfQuestions());
			$worksheet->setCell($row, $col++, ($userdata->getQuestionsWorkedThroughInPercent()) . '%');
			
			// getTimeOfWork() was renamed to getTimeOnTask() in ILIAS 10
			$time_of_work = $userdata->getTimeOnTask();
			$worksheet->setCell($row, $col++, $this->secondsToTimeString($time_of_work));
			$worksheet->setCell($row, $col++, $this->secondsToTimeString(
				$userdata->getQuestionsWorkedThrough() ? (int) ($time_of_work / $userdata->getQuestionsWorkedThrough()) : 0
			));
			// getFirstVisit()/getLastVisit() return ?DateTimeImmutable since ILIAS 10
			$worksheet->setCell($row, $col++, $this->toIlDateTime($userdata->getFirstVisit()));
			$worksheet->setCell($row, $col++, $this->toIlDateTime($userdata->getLastVisit()));
			
			$median = $statistics->median() ?? 0.0;
			$pct = $userdata->getMaxpoints() ? $median / $userdata->getMaxpoints() * 100.0 : 0;

            $mark = $this->test_obj->getMarkSchema()->getMatchingMark((float) $pct);
			$mark_short_name = "";
			
			if(is_object($mark))
			{
				$mark_short_name = $mark->getShortName();
			}
			
			$worksheet->setCell($row, $col++, $mark_short_name);
			$worksheet->setCell($row, $col++, $statistics->rank($userdata->getReached()) ?? '');
			// rank_median() was renamed to rankMedian() in ILIAS 10
			$worksheet->setCell($row, $col++, $statistics->rankMedian() ?? '');
			$worksheet->setCell($row, $col++, $statistics->count());
			$worksheet->setCell($row, $col++, $median);
			
			if($this->test_obj->getPassScoring() == ilObjTest::SCORE_BEST_PASS)
			{
				$worksheet->setCell($row, $col++, $userdata->getBestPass() + 1);
			}
			else
			{
				$worksheet->setCell($row, $col++, $userdata->getLastPass() + 1);
			}
			
			$startcol = $col;
			
			for($pass = 0; $pass <= $userdata->getLastPass(); $pass++)
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
					if(is_array($userdata->getQuestions($pass)))
					{
						$evaluatedQuestions = $userdata->getQuestions($pass);
						
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
							$question_data = $userdata->getPass($pass)?->getAnsweredQuestionByQuestionId((int) $question["id"]);
                            if ($question_data && $question_data["reached"]) {
                                $worksheet->setCell($row, $col, $question_data["reached"]);
                            }
							if($this->test_obj->isRandomTest())
							{
								// random test requires question headers for every participant
								// and we allready skipped a row for that reason ( --> row - 1)
								$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col) . ($row - 1),  preg_replace("/<.*?>/", "", $data->getQuestionTitle((int) $question["id"])));
							}
							else
							{
								if($pass == 0 && !$firstrowwritten)
								{
									$this->setFormattedExcelTitle($worksheet, $worksheet->getColumnCoord($col) . 1, $data->getQuestionTitle((int) $question["id"]));
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

	/**
	 * Replacement for ilAssExcelFormatHelper::setFormattedExcelTitle(), which was removed in ILIAS 10.
	 */
	private function setFormattedExcelTitle(ilExcel $worksheet, string $coordinates, string $value): void
	{
		$worksheet->setCellByCoordinates($coordinates, $value);
		$worksheet->setColors($coordinates, self::EXCEL_BACKGROUND_COLOR);
		$worksheet->setBold($coordinates);
	}

	/**
	 * ilTestEvaluationUserData::$mark is a typed, non nullable property that stays
	 * uninitialised when no mark step matches, so it cannot be read unguarded.
	 */
	private function getMarkShortName(ilTestEvaluationUserData $userdata): string
	{
		try {
			return $userdata->getMark()->getShortName();
		} catch (Error $e) {
			return '';
		}
	}

	private function secondsToTimeString(int $seconds): string
	{
		$hours = (int) floor($seconds / 3600);
		$seconds -= $hours * 3600;
		$minutes = (int) floor($seconds / 60);
		$seconds -= $minutes * 60;
		return sprintf("%02d:%02d:%02d", $hours, $minutes, $seconds);
	}

	private function toIlDateTime(?DateTimeImmutable $date_time): ilDateTime|string
	{
		if ($date_time === null) {
			return '';
		}
		return new ilDateTime($date_time->getTimestamp(), IL_CAL_UNIX);
	}
}
