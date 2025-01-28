import React, { useState } from 'react';
import { Card, CardHeader, CardTitle, CardContent } from './ui/card';
import { Button } from './ui/button';
import { Input } from './ui/input';
import { Loader2 } from 'lucide-react';
import { Alert, AlertDescription } from './ui/alert';

const ExcelAutomationForm = () => {
    const [formData, setFormData] = useState({
        version: '',
        filePathUrl: '',
        Destination_filePathUrl: '', // Changed to match the exact case
        language_name: [  // Changed from 'languages' to 'language_name'
            "English (United States)",
            "French (France)",
            "Italian (Italy)",
            "Russian (Russia)",
            "Japanese (Japan)",
            "Chinese (Simplified, China)",
            "Chinese (Traditional, Taiwan)",
            "Arabic (Oman)"
        ]
    });

    const [loading, setLoading] = useState(false);
    const [message, setMessage] = useState('');
    const [error, setError] = useState('');

    const handleInputChange = (e) => {
        const { name, value } = e.target;
        setFormData(prev => ({
            ...prev,
            [name]: value
        }));
    };

    const handleLanguageToggle = (language) => {
        setFormData(prev => ({
            ...prev,
            language_name: prev.language_name.includes(language)  // Changed to language_name
                ? prev.language_name.filter(lang => lang !== language)
                : [...prev.language_name, language]
        }));
    };

    const handleSubmit = async (e) => {
        e.preventDefault();
        setLoading(true);
        setMessage('');
        setError('');

        try {
            const response = await fetch('http://localhost:8080/excel/automate', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(formData)
            });

            if (!response.ok) {
                throw new Error('Failed to process the request');
            }

            const data = await response.text();
            setMessage(data);
        } catch (err) {
            setError(err.message);
        } finally {
            setLoading(false);
        }
    };

    const allLanguages = [
        "English (United States)",
        "French (France)",
        "Italian (Italy)",
        "Russian (Russia)",
        "Japanese (Japan)",
        "Chinese (Simplified, China)",
        "Chinese (Traditional, Taiwan)",
        "Arabic (Oman)"
    ];

    return (
        <div className="container mx-auto p-4 max-w-2xl">
            <Card>
                <CardHeader>
                    <CardTitle>Excel Automation Form</CardTitle>
                </CardHeader>
                <CardContent>
                    <form onSubmit={handleSubmit} className="space-y-4">
                        <div>
                            <label className="block text-sm font-medium mb-1">Version</label>
                            <Input
                                name="version"
                                value={formData.version}
                                onChange={handleInputChange}
                                placeholder="e.g., V3_01_17_01"
                                required
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium mb-1">File Path URL</label>
                            <Input
                                name="filePathUrl"
                                value={formData.filePathUrl}
                                onChange={handleInputChange}
                                placeholder="C:\PowerAutomate\X2_Collect_V3_01_17_01_00_.xlsm"
                                required
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium mb-1">Destination Path URL</label>
                            <Input
                                name="Destination_filePathUrl"  // Changed to match the exact case
                                value={formData.Destination_filePathUrl}
                                onChange={handleInputChange}
                                placeholder="C:\PowerAutomate\"
                                required
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium mb-2">Languages</label>
                            <div className="grid grid-cols-2 gap-2">
                                {allLanguages.map(language => (
                                    <label key={language} className="flex items-center space-x-2">
                                        <input
                                            type="checkbox"
                                            checked={formData.language_name.includes(language)}  // Changed to language_name
                                            onChange={() => handleLanguageToggle(language)}
                                            className="rounded border-gray-300"
                                        />
                                        <span className="text-sm">{language}</span>
                                    </label>
                                ))}
                            </div>
                        </div>

                        <Button
                            type="submit"
                            className="w-full"
                            disabled={loading}
                        >
                            {loading ? (
                                <>
                                    <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                                    Processing...
                                </>
                            ) : 'Submit'}
                        </Button>

                        {message && (
                            <Alert className="mt-4 bg-green-50">
                                <AlertDescription>{message}</AlertDescription>
                            </Alert>
                        )}

                        {error && (
                            <Alert className="mt-4 bg-red-50">
                                <AlertDescription className="text-red-600">{error}</AlertDescription>
                            </Alert>
                        )}
                    </form>
                </CardContent>
            </Card>
        </div>
    );
};

export default ExcelAutomationForm;