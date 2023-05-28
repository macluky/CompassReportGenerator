from reportgenerator import ModelData, ResultData, ReportGenerator
from questionsparser import QuestionsParser
from adviceparser import AdviceParser
from weightparser import WeightParser
from dynamicweightparser import DynamicWeightParser
from typeformresultsparser import TypeFormResultsParser
import helpers


def process_results_for_model(results, model, use_dynamic_weights=False):
    qp = QuestionsParser(_count=model.question_count, _levels=model.model_layers_count, _file=model.file)
    ap = AdviceParser(_count=model.recommendations, _max_model_layers=model.model_layers_count, _file=model.file)
    wp = WeightParser(_file=model.file, _weight_tab=model.weight_tab)
    tp = TypeFormResultsParser(_questions_parser=qp, _count=model.question_count, _file=results.file,
                               _nr_candidates=results.nr_of_candidates)

    # do we need to average all scores first and use as weights?
    if use_dynamic_weights:
        print("Warning: using average scores as weight model, fixed model ignored")
        dwp = DynamicWeightParser(wp, tp.candidates)
        wp = dwp

    all_answers = []
    # TODO: recursively create separate generators. This can be split off in a separate file
    for i in range(0, results.nr_of_candidates):
        report = ReportGenerator(_results_parser=tp, _question_parser=qp, _advice_parser=ap,
                                 _weight_parser=wp)
        report.select_candidate(i)
        report.set_name()

        # save data for calculating averages later on
        all_answers.append(report.results)

        # shouldn't we pass our parsers instead of addressing them as globals?
        report.score_analysis()
        report.add_graphs_with_ref_level()
        report.add_recommendations()
        report.save(results.result_path)

    # create an average report for all candidates
    if not use_dynamic_weights:
        report = ReportGenerator(_results_parser=tp, _question_parser=qp, _advice_parser=ap,
                                 _weight_parser=wp)
        report.name = m.weight_tab
        report.email = m.weight_tab
        report.results = helpers.average_all_results(all_answers)
        report.score_analysis()
        report.add_graphs_with_ref_level()
        report.add_recommendations()
        report.save(results.result_path)

# main
rp = "/Users/macluky/Documents/Work/Expandior/Clients/TopDesk/Assessment"

m = ModelData(_weight_tab="TOPdesk PO")
r = ResultData(_nr_of_candidates=12, _result_path=rp, _result_db="responses-po.xlsx")
g = ReportGenerator()

process_results_for_model(r, m, False)

m = ModelData(_weight_tab="TOPdesk PM")
r = ResultData(_nr_of_candidates=2, _result_path=rp, _result_db="responses-pm.xlsx")
process_results_for_model(r, m, False)