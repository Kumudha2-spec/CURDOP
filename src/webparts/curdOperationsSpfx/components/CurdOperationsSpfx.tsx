import * as React from 'react';
import {
  Box,
  Button,
  FormControl,
  InputLabel,
  MenuItem,
  Select,
  TextField,
  Typography,
  Paper,
  Radio,
  RadioGroup,
  FormControlLabel,
  FormLabel,
  Alert,
  Stack
} from '@mui/material';
import type { SelectChangeEvent } from '@mui/material';

import type { ICurdOperationsSpfxProps } from './ICurdOperationsSpfxProps';

export interface IEmployeeFeedbackState {
  employeeName: string;
  email: string;
  department: string;
  feedback: string;
  rating: string;
  isSubmitted: boolean;
  errorMessage: string;
}

export default class CurdOperationsSpfx extends React.Component<
  ICurdOperationsSpfxProps,
  IEmployeeFeedbackState
> {
  constructor(props: ICurdOperationsSpfxProps) {
    super(props);
    this.state = {
      employeeName: '',
      email: '',
      department: '',
      feedback: '',
      rating: '3',
      isSubmitted: false,
      errorMessage: ''
    };
  }

  private handleTextChange = (field: keyof IEmployeeFeedbackState) => (
    event: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>
  ): void => {
    this.setState({
      ...this.state,
      [field]: event.target.value,
      isSubmitted: false,
      errorMessage: ''
    });
  };

  private handleDepartmentChange = (event: SelectChangeEvent): void => {
    this.setState({
      department: event.target.value,
      isSubmitted: false,
      errorMessage: ''
    });
  };

  private handleSubmit = (event: React.FormEvent): void => {
    event.preventDefault();

    const { employeeName, email, department, feedback, rating } = this.state;

    // Basic validation
    if (!employeeName.trim() || !email.trim() || !department || !feedback.trim()) {
      this.setState({ errorMessage: 'Please fill in all mandatory fields.' });
      return;
    }

    console.log('Feedback submitted:', {
      employeeName,
      email,
      department,
      feedback,
      rating
    });

    // TODO: Invoke SharePoint REST API / PnPjs to create the list item here

    this.setState({
      isSubmitted: true,
      errorMessage: ''
    });
  };

  private handleReset = (): void => {
    this.setState({
      employeeName: '',
      email: '',
      department: '',
      feedback: '',
      rating: '3',
      isSubmitted: false,
      errorMessage: ''
    });
  };

  public render(): React.ReactElement<ICurdOperationsSpfxProps> {
    const { employeeName, email, department, feedback, rating, isSubmitted, errorMessage } =
      this.state;

    return (
      <Box sx={{ maxWidth: 700, margin: '30px auto' }}>
        <Paper elevation={3} sx={{ padding: 4, borderRadius: 2 }}>
          <Typography variant="h4" component="h1" gutterBottom color="primary">
            Employee Feedback
          </Typography>

          <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
            Please share your feedback with us. All fields marked with * are required.
          </Typography>

          {isSubmitted && (
            <Alert severity="success" sx={{ mb: 3 }}>
              Thank you! Your feedback has been recorded successfully.
            </Alert>
          )}

          {errorMessage && (
            <Alert severity="error" sx={{ mb: 3 }}>
              {errorMessage}
            </Alert>
          )}

          <Box component="form" onSubmit={this.handleSubmit} noValidate>
            <TextField
              fullWidth
              required
              label="Employee Name"
              margin="normal"
              value={employeeName}
              onChange={this.handleTextChange('employeeName')}
            />

            <TextField
              fullWidth
              required
              label="Email"
              type="email"
              margin="normal"
              value={email}
              onChange={this.handleTextChange('email')}
            />

            <FormControl fullWidth margin="normal" required>
              <InputLabel id="dept-select-label">Department</InputLabel>
              <Select
                labelId="dept-select-label"
                label="Department"
                value={department}
                onChange={this.handleDepartmentChange}
              >
                <MenuItem value="IT">IT</MenuItem>
                <MenuItem value="HR">HR</MenuItem>
                <MenuItem value="Finance">Finance</MenuItem>
                <MenuItem value="Marketing">Marketing</MenuItem>
              </Select>
            </FormControl>

            <TextField
              fullWidth
              required
              label="Feedback"
              multiline
              rows={4}
              margin="normal"
              value={feedback}
              onChange={this.handleTextChange('feedback')}
            />

            <FormControl component="fieldset" sx={{ mt: 3, display: 'block' }}>
              <FormLabel component="legend">Rating (1 to 5)</FormLabel>
              <RadioGroup
                row
                value={rating}
                onChange={this.handleTextChange('rating')}
              >
                <FormControlLabel value="1" control={<Radio />} label="1" />
                <FormControlLabel value="2" control={<Radio />} label="2" />
                <FormControlLabel value="3" control={<Radio />} label="3" />
                <FormControlLabel value="4" control={<Radio />} label="4" />
                <FormControlLabel value="5" control={<Radio />} label="5" />
              </RadioGroup>
            </FormControl>

            <Stack direction="row" spacing={2} sx={{ mt: 3 }}>
              <Button
                variant="outlined"
                color="secondary"
                onClick={this.handleReset}
                sx={{ minWidth: 120 }}
              >
                Clear
              </Button>
              <Button
                type="submit"
                variant="contained"
                color="primary"
                fullWidth
              >
                Submit Feedback
              </Button>
            </Stack>
          </Box>
        </Paper>
      </Box>
    );
  }
}