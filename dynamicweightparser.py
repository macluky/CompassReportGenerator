import helpers


class DynamicWeightParser:

    def __init__(self, wp, candidates):
        keys = wp.weights.keys()
        self.weights = dict()

        for key in keys:
            p1 = key[0]
            p2 = key[1]
            scores = []
            for candidate in candidates:
                #print("averaging: "+candidate.name)
                score = helpers.calc_score_for_label(candidate.answers, p2)
                scores.append(score)
            average = helpers.average_of_scores(scores)
            self.weights[key] = average

    def weight_label(self, label):
        # none is level 0
        if label is None:
            total = 0
            for w in self.weights.values():
                total += w
            return total/6
        else:
            count = 0
            total = 0
            for t in self.weights.keys():
                if label in t:
                    count += 1
                    total += self.weights[t]
            return total/count


