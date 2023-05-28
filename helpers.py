def all_labels_at_level(results, depth):
    labels = []
    for result in results:
        label = result.levels[depth]
        if label in labels:
            continue
        labels.append(label)

    return labels


def average_of_scores(list_of_scores):
    total = 0
    for score in list_of_scores:
        total += score
    return total / len(list_of_scores)


def calc_score_for_label(results, label):
    # assumes unique label at each level!
    scores = []

    for result in results:
        if label in result.levels:
            scores.append(result.score)

    if len(scores) > 0:
        return average_of_scores(scores)
    else:
        print("No score for " + label)
        return None


def depth_of_label(results_parser, results, label):
    for depth in range(0, results_parser.questions_parser.levels):
        labels_at_level = all_labels_at_level(results, depth)
        if label in labels_at_level:
            return depth
    print("Error label not found: " + label)
    return None


def sub_labels_of_label(results_parser, results, label):
    depth = depth_of_label(results_parser, results, label)
    sub_labels = []
    for result in results:
        if (depth + 1) >= len(result.levels):
            #there are no sublevels
            return None
        sub_label = result.levels[depth + 1]
        parent = result.levels[depth]
        if parent == label:
            if sub_label in sub_labels:
                continue
            else:
                sub_labels.append(sub_label)
    return sub_labels


def print_scores_at_level(depth):
    labels = all_labels_at_level(depth)
    for label in labels:
        score = calc_score_for_label(label)
        print("Score " + str(score) + " for " + label)


def score_at_level(results, level):
    labels = all_labels_at_level(results, level)
    # average for our level
    score = 0
    for label in labels:
        score += calc_score_for_label(results, label)
    score /= len(labels)
    return score


def weight_for_parent_label(weight_parser, results, label):
    for result in results:
        labels = result.levels
        if label in labels:
            parent = labels[1]
            # print("parent of " + label + " is " + parent + " weight " + str(wp.weight_label(parent)))
            return weight_parser.weight_label(parent)


def substitute_image_placeholder(paragraph, image_filename):
    # --- start with removing the placeholder text ---
    paragraph.text = ""
    # paragraph.text = paragraph.text[1:len(paragraph.text) - 1]
    # --- then append a run containing the image ---
    run = paragraph.add_run()
    run.add_picture(image_filename)


def average_all_results(all_assessments):
    averages = []
    # voor alle vragen, voor alle kandidaten
    for i in range(0, len(all_assessments[0])):
        scores = []
        for assesment in all_assessments:
            score = assesment[i].score
            scores.append(score)
        question = all_assessments[0][i]
        question.score = average_of_scores(scores)
        averages.append(question)
    return averages







