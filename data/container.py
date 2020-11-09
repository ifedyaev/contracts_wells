class BorderStatistWells:
    """
    class contains statistic by border contracts
    """

    def __init__(self):
        self.count_days = 0
        self.count_wells = 0
        self.sum_costs = 0.0
        self.costs_day = 0.0

    def compute_costs_day(self) -> None:
        """
        compute costs days
        :return:
        """
        if self.costs_day == 0:
            self.costs_day = 0.0
        else:
            self.costs_day = self.sum_costs / self.count_days
        return
