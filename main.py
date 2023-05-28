from reportgenerator import ModelData, ResultData, ReportGenerator
from questionsparser import QuestionsParser
from adviceparser import AdviceParser
from weightparser import WeightParser
from dynamicweightparser import DynamicWeightParser
from typeformresultsparser import TypeFormResultsParser
import helpers


def process_results_for_model(results, model, template_path, use_dynamic_weights=False):
    qp = QuestionsParser(_count=model.question_count, _levels=model.model_layers_count, _file=model.file)
    ap = AdviceParser(_count=model.recommendations, _max_model_layers=model.model_layers_count, _file=model.file)
    wp = WeightParser(_file=model.file, _weight_tab=model.weight_tab, _count=model.area_count)
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
                                 _weight_parser=wp, _file=template_path)
        report.select_candidate(i)
        report.set_name()

        # save data for calculating averages later on
        all_answers.append(report.results)

        # TODO: shouldn't we pass our parsers instead of addressing them as globals?
        # only deepdive on level 3 models
        if model.model_layers_count > 2:
            report.score_analysis()
        report.add_graphs_with_ref_level()
        report.add_recommendations()
        report.save(results.result_path)

    # create an average report for all candidates
    if not use_dynamic_weights:
        report = ReportGenerator(_results_parser=tp, _question_parser=qp, _advice_parser=ap,
                                 _weight_parser=wp, _file=template_path)
        report.name = model.weight_tab
        report.email = model.weight_tab
        report.results = helpers.average_all_results(all_answers)
        # only deepdive on level 3 models
        if model.model_layers_count > 2:
            report.score_analysis()
        report.add_graphs_with_ref_level()
        report.add_recommendations()
        report.save(results.result_path)


# Based on model 4.9
def run_1():
    rp = "/Users/macluky/Documents/Work/Expandior/Clients/TopDesk/Assessment"
    m = ModelData(_weight_tab="TOPdesk PO")
    r = ResultData(_nr_of_candidates=12, _result_path=rp, _result_db="responses-po.xlsx")
    g = ReportGenerator()
    process_results_for_model(r, m, g.file, False)
    m = ModelData(_weight_tab="TOPdesk PM")
    r = ResultData(_nr_of_candidates=2, _result_path=rp, _result_db="responses-pm.xlsx")
    process_results_for_model(r, m, False)


# Based on model 5.0
def run_2():
    rp = "/Users/macluky/Library/CloudStorage/OneDrive-SharedLibraries-ExpandiorAcademyB.V/Expandior Team - Documents/Product/Compass (Maturity Scan Personal)/Release 5.0/Resources"

    # you can drop to 67 questions when partnering and product knowledge are ommited (also from the results!!!)
    m = ModelData(_question_count=74, _model_layers_count=2, _model_area_count=18, _question_path=rp, _question_db="PPC Question DB v5.0 rc2.xlsx")
    r = ResultData(_nr_of_candidates=1, _result_path=rp, _result_db="test_responses.xlsx")
    process_results_for_model(r, m, rp+"/ReportTemplate.docx", False)

# main
run_2()
#check backwards compatibilty
#run_1()