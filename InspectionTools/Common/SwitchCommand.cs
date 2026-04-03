namespace InspectionTools.Common {
    internal record SwitchCommand {
        public DcsMode DcsMode { get; init; }
        public DmmMode DmmMode { get; init; }
        public string Text { get; init; } = string.Empty;
        public string Adc { get; init; } = string.Empty;
        public string Visa { get; init; } = string.Empty;
        public string Gpib { get; init; } = string.Empty;
        public bool Query { get; init; } = false;
    }
}
